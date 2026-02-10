# -*- coding: utf-8 -*-
import sys
import os
import time
import ctypes
import threading
import pyautogui
import pyperclip
import traceback
from PIL import Image

# Check for psutil
try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False

# Windows API
GetAsyncKeyState = ctypes.windll.user32.GetAsyncKeyState
try:
    GetCurrentProcessorNumber = ctypes.windll.kernel32.GetCurrentProcessorNumber
    GetCurrentProcessorNumber.restype = ctypes.c_ulong
    HAS_KERNEL_CPU = True
except:
    HAS_KERNEL_CPU = False

pyautogui.FAILSAFE = False 
pyautogui.PAUSE = 0

from .utils import write_log
from .config import GLOBAL_CONFIG

# --------------------------
# 独立看门狗线程
# --------------------------
class FailsafeWatchdog(threading.Thread):
    def __init__(self, engine):
        super().__init__()
        self.engine = engine
        self.daemon = True 
        self.running = True

    def run(self):
        write_log(">>> 看门狗线程启动")
        while self.running:
            try:
                if self.engine.enable_key_stop:
                    if GetAsyncKeyState(0x1B) & 0x8000: 
                        self.trigger_stop("用户按下了【ESC键】")
                        return
                    if GetAsyncKeyState(0x04) & 0x8000: 
                        self.trigger_stop("用户按下了【鼠标中键】")
                        return

                if self.engine.enable_tr_stop:
                    x, y = pyautogui.position()
                    w, h = pyautogui.size()
                    if x > (w - 10) and y < 10:
                        self.trigger_stop("检测到鼠标【右上角急停】")
                        return

                if self.engine.enable_tm_stop:
                    if int(time.time() * 100) % 10 == 0: 
                        hwnd = ctypes.windll.user32.GetForegroundWindow()
                        length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
                        if length > 0:
                            buff = ctypes.create_unicode_buffer(length + 1)
                            ctypes.windll.user32.GetWindowTextW(hwnd, buff, length + 1)
                            # buff.value guarantees a string 
                            window_title = buff.value
                            if "任务管理器" in window_title or "Task Manager" in window_title:
                                self.trigger_stop("检测到【任务管理器】前台")
                                return
                time.sleep(0.02)
            except Exception as e:
                # write_log(f"Watchdog error: {e}")
                time.sleep(1)

    def trigger_stop(self, reason):
        if not self.engine.stop_requested:
            write_log(f">>> 看门狗触发: {reason}")
            self.engine.log(f"!!! {reason} -> 停止 !!!")
            self.engine.stop() 
            try: ctypes.windll.user32.MessageBeep(0xFFFFFFFF)
            except: pass

    def kill(self):
        self.running = False

# --------------------------
# 核心引擎 (V45+ 内核)
# --------------------------
class RPAEngine:
    def __init__(self):
        self.is_running = False
        self.stop_requested = False
        
        self.min_scale = 1.0
        self.max_scale = 1.0
        self.confidence = 0.8
        self.scan_region = None 
        
        self.dodge_x1 = 100
        self.dodge_y1 = 100
        self.dodge_x2 = 200
        self.dodge_y2 = 100
        self.enable_dodge = False
        self.enable_double_dodge = False
        self.double_dodge_wait = 0.015
        
        self.move_duration = 0.0
        self.click_hold = 0.04
        self.settlement_wait = 0.0
        self.timeout_val = 0.0
        
        self.enable_tm_stop = True 
        self.enable_tr_stop = True 
        self.enable_key_stop = True
        
        self.callback_msg = None
        self.opencv_available = False 
        self.img_cache = {} 
        self.scaled_templates_cache = {}

        self.check_engine_status()
        self.set_high_priority()

    def set_high_priority(self):
        try:
            pid = os.getpid()
            handle = ctypes.windll.kernel32.OpenProcess(0x0100, True, pid)
            ctypes.windll.kernel32.SetPriorityClass(handle, 0x00000080)
        except: pass

    def check_engine_status(self):
        try:
            import cv2
            import numpy
            # Simple check if imported modules are working
            img = numpy.zeros((10, 10, 3), dtype=numpy.uint8)
            cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
            self.opencv_available = True
            write_log("OpenCV/NumPy 引擎就绪。")
        except:
            self.opencv_available = False
            write_log("OpenCV 引擎不可用。")

    def stop(self):
        self.stop_requested = True
        self.is_running = False

    def log(self, msg):
        write_log(msg)
        if self.callback_msg: self.callback_msg(msg)

    def check_stop_flag(self):
        return self.stop_requested

    def load_and_precompute(self, tasks):
        if not self.opencv_available: return
        try:
            import cv2
            import numpy as np
            
            write_log("正在预加载资源...")
            for task in tasks:
                path = str(task.get("value", ""))
                if not path or not os.path.exists(path): continue
                # Types that involve images: 1.0 (click), 2.0 (double click), 3.0 (right click), 8.0 (hover)
                # But actually find_target handles finding the image.
                if task.get("type") not in [1.0, 2.0, 3.0, 8.0]: continue
                
                img = Image.open(path)
                img.load()
                self.img_cache[path] = img
                
                if self.min_scale != 1.0 or self.max_scale != 1.0:
                    if img.mode != 'L': img = img.convert('L')
                    template = np.array(img)
                    
                    templates_list = []
                    steps = int((self.max_scale - self.min_scale) / 0.05) + 1
                    # Avoid division by zero if steps is weird, but it should be fine.
                    # linspace handles it.
                    if steps < 1: steps = 1
                    
                    for scale in np.linspace(self.min_scale, self.max_scale, steps):
                        if 0.99 < scale < 1.01: continue
                        rw = int(template.shape[1] * scale)
                        rh = int(template.shape[0] * scale)
                        if rw < 1 or rh < 1: continue
                        resized_tpl = cv2.resize(template, (rw, rh))
                        templates_list.append(resized_tpl)
                    
                    self.scaled_templates_cache[path] = templates_list
            write_log("资源预加载完成。")
        except Exception as e:
            write_log(f"预计算失败: {e}")

    def find_target_optimized(self, img_path):
        try:
            screenshot_pil = pyautogui.screenshot(region=self.scan_region)
        except: return None
        
        offset_x = self.scan_region[0] if self.scan_region else 0
        offset_y = self.scan_region[1] if self.scan_region else 0

        if not self.opencv_available:
            if img_path in self.img_cache:
                try: 
                    res = pyautogui.locate(self.img_cache[img_path], screenshot_pil, confidence=self.confidence)
                    if res:
                        cx = res.left + (res.width / 2) + offset_x
                        cy = res.top + (res.height / 2) + offset_y
                        return (cx, cy)
                except: pass
            elif os.path.exists(img_path):
                 try:
                    res = pyautogui.locate(img_path, screenshot_pil, confidence=self.confidence)
                    if res:
                        cx = res.left + (res.width / 2) + offset_x
                        cy = res.top + (res.height / 2) + offset_y
                        return (cx, cy)
                 except: pass
            return None

        # OpenCV logic
        import cv2
        import numpy as np
        
        screen_np = np.array(screenshot_pil)
        # Convert RGB to GRAY
        # pyautogui returns RGB usually (PIL image)
        screen_gray = cv2.cvtColor(screen_np, cv2.COLOR_RGB2GRAY)
        
        if img_path not in self.img_cache:
            if os.path.exists(img_path):
                try:
                    img = Image.open(img_path)
                    img.load()
                    self.img_cache[img_path] = img
                except: return None
            else:
                return None
        
        pil_template = self.img_cache[img_path]
        
        try:
            if pil_template.mode != 'L': pil_template = pil_template.convert('L')
            tpl_gray = np.array(pil_template)
            
            if tpl_gray.shape[0] > screen_gray.shape[0] or tpl_gray.shape[1] > screen_gray.shape[1]:
                pass 
            else:
                res = cv2.matchTemplate(screen_gray, tpl_gray, cv2.TM_CCOEFF_NORMED)
                min_v, max_v, min_l, max_l = cv2.minMaxLoc(res)
                if max_v >= self.confidence:
                    h, w = tpl_gray.shape[:2]
                    final_x = max_l[0] + w//2 + offset_x
                    final_y = max_l[1] + h//2 + offset_y
                    return (final_x, final_y)
        except: pass
        
        # Checked cached scaled templates
        if img_path in self.scaled_templates_cache:
            for resized_tpl in self.scaled_templates_cache[img_path]:
                if self.check_stop_flag(): return None
                try:
                    if resized_tpl.shape[0] > screen_gray.shape[0] or resized_tpl.shape[1] > screen_gray.shape[1]:
                        continue
                    res = cv2.matchTemplate(screen_gray, resized_tpl, cv2.TM_CCOEFF_NORMED)
                    min_v, max_v, min_l, max_l = cv2.minMaxLoc(res)
                    if max_v >= self.confidence:
                        h, w = resized_tpl.shape[:2]
                        final_x = max_l[0] + w//2 + offset_x
                        final_y = max_l[1] + h//2 + offset_y
                        return (final_x, final_y)
                except: continue
        
        return None

    def mouseClick(self, clickTimes, lOrR, img_path, reTry):
        start_time = time.time()
        
        _move = self.move_duration
        _hold = self.click_hold
        _dodge_en = self.enable_dodge
        _dx1, _dy1 = self.dodge_x1, self.dodge_y1
        _dx2, _dy2 = self.dodge_x2, self.dodge_y2
        _dbl_dodge = self.enable_double_dodge
        _dbl_wait = self.double_dodge_wait
        _timeout = self.timeout_val
        _settle = self.settlement_wait
        
        while True:
            if self.check_stop_flag(): return
            if _timeout > 0.001 and (time.time() - start_time > _timeout): return

            location_tuple = self.find_target_optimized(img_path)

            if location_tuple:
                try:
                    x, y = location_tuple
                    
                    pyautogui.moveTo(x, y, duration=_move)
                    for _ in range(clickTimes):
                        pyautogui.mouseDown(button=lOrR)
                        time.sleep(_hold)
                        pyautogui.mouseUp(button=lOrR)
                        if clickTimes > 1: time.sleep(0.02)
                    
                    if _settle > 0: time.sleep(_settle)
                    
                    if _dodge_en:
                        pyautogui.moveTo(_dx1, _dy1, duration=0)
                        if _dbl_dodge:
                            time.sleep(_dbl_wait) 
                            pyautogui.moveTo(_dx2, _dy2, duration=0)
                    
                except Exception as e: self.log(f"Err: {e}")
                
                if reTry != -1: return
                else:
                    time.sleep(0.01)
                    continue
            
            if _timeout <= 0.001: return 
            time.sleep(0.001) 

    def run_tasks(self, tasks, loop_forever=False, callback_msg=None):
        self.is_running = True
        self.stop_requested = False
        self.callback_msg = callback_msg
        
        self.img_cache = {}
        self.scaled_templates_cache = {}
        self.load_and_precompute(tasks)
        
        if self.scan_region:
            write_log(f"区域模式: {self.scan_region}")
        
        try:
            while True:
                for idx, task in enumerate(tasks):
                    if self.check_stop_flag():
                        if callback_msg: callback_msg("任务由看门狗终止")
                        return

                    cmd = task.get("type")
                    val = task.get("value")
                    retry = task.get("retry", 1)
                    
                    if cmd == 1.0: self.mouseClick(1, "left", val, retry)
                    elif cmd == 2.0: self.mouseClick(2, "left", val, retry)
                    elif cmd == 3.0: self.mouseClick(1, "right", val, retry)
                    elif cmd == 8.0:
                        loc = self.find_target_optimized(val)
                        if loc: pyautogui.moveTo(loc[0], loc[1], duration=self.move_duration)
                    elif cmd == 4.0: 
                        pyperclip.copy(str(val)); pyautogui.hotkey('ctrl', 'v'); time.sleep(0.2)
                    elif cmd == 5.0: 
                        t_end = time.time() + float(val)
                        while time.time() < t_end:
                            if self.check_stop_flag(): return
                            time.sleep(0.05)
                    elif cmd == 6.0: pyautogui.scroll(int(val))
                    elif cmd == 7.0: pyautogui.hotkey(*[k.strip() for k in str(val).lower().split('+')])
                    elif cmd == 9.0:
                        path = str(val)
                        if os.path.isdir(path): path = os.path.join(path, time.strftime("ss_%H%M%S.png"))
                        try: pyautogui.screenshot(path, region=self.scan_region)
                        except: pass

                if not loop_forever: break
                if self.check_stop_flag(): return
                
        except Exception as e:
            self.log(f"引擎异常: {e}")
        finally:
            self.is_running = False
            if callback_msg: callback_msg("结束")

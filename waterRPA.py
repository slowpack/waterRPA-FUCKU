import sys
import os
import time
import json
import traceback
import ctypes
import threading

# ---------------------------------------------------------
# 核心库导入
# ---------------------------------------------------------
try:
    os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "0"
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    ctypes.windll.shcore.SetProcessDpiAwareness(1) 
except:
    try: ctypes.windll.user32.SetProcessDPIAware()
    except: pass

from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                               QPushButton, QLabel, QComboBox, QLineEdit, QScrollArea, 
                               QFileDialog, QTextEdit, QMessageBox, QFrame, QCheckBox, QGroupBox, QToolTip,
                               QListWidget, QListWidgetItem, QAbstractItemView, QRubberBand)
from PySide6.QtCore import Qt, QThread, Signal, QTimer, QSize, QRect, QSettings, QPoint
from PySide6.QtGui import QCursor, QFont, QColor, QPalette, QBrush, QPen, QPainter, QRegion
import pyperclip
from PIL import Image
import pyautogui

try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False

GetAsyncKeyState = ctypes.windll.user32.GetAsyncKeyState
try:
    GetCurrentProcessorNumber = ctypes.windll.kernel32.GetCurrentProcessorNumber
    GetCurrentProcessorNumber.restype = ctypes.c_ulong
    HAS_KERNEL_CPU = True
except:
    HAS_KERNEL_CPU = False

pyautogui.FAILSAFE = False 
pyautogui.PAUSE = 0

# ---------------------------------------------------------
# 全局配置
# ---------------------------------------------------------
GLOBAL_CONFIG = {
    "log_to_file": False,
    "log_to_ui": True
}

def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

def get_log_path():
    return os.path.join(get_base_dir(), "rpa_debug_log.txt")

def write_log(msg):
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    formatted_msg = f"[{timestamp}] {msg}"
    if GLOBAL_CONFIG["log_to_file"]:
        try:
            with open(get_log_path(), "a", encoding="utf-8") as f:
                f.write(formatted_msg + "\n")
        except: pass

def global_exception_handler(exctype, value, tb):
    err_msg = "".join(traceback.format_exception(exctype, value, tb))
    write_log(f"!!! 严重崩溃 !!! {value}\n{err_msg}")
    sys.__excepthook__(exctype, value, tb)

sys.excepthook = global_exception_handler

# --------------------------
# 区域选择窗口
# --------------------------
class RegionWindow(QWidget):
    region_selected = Signal(tuple) # x, y, w, h

    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_DeleteOnClose)
        
        self.setCursor(Qt.CrossCursor)
        self.setMouseTracking(True)
        
        virtual_rect = QApplication.primaryScreen().virtualGeometry()
        self.setGeometry(virtual_rect)
        
        phys_w, phys_h = pyautogui.size()
        log_w = virtual_rect.width()
        log_h = virtual_rect.height()
        self.scale_x = phys_w / log_w
        self.scale_y = phys_h / log_h
        
        self.start_point = None
        self.end_point = None
        self.current_pos = QPoint(0, 0)
        self.selection_rect = QRect()
        
        self.show()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        bg_color = QColor(0, 0, 0, 100) 
        
        if self.selection_rect.isValid():
            mask_region = QRegion(self.rect())
            selection_region = QRegion(self.selection_rect)
            overlay_region = mask_region.subtracted(selection_region)
            
            painter.setClipRegion(overlay_region)
            painter.fillRect(self.rect(), bg_color)
            
            painter.setClipping(False)
            pen = QPen(QColor(0, 255, 0), 2)
            pen.setStyle(Qt.DashLine)
            painter.setPen(pen)
            painter.setBrush(Qt.NoBrush)
            painter.drawRect(self.selection_rect)
            
            real_w = int(self.selection_rect.width() * self.scale_x)
            real_h = int(self.selection_rect.height() * self.scale_y)
            info_text = f"选区:{self.selection_rect.width()}x{self.selection_rect.height()} (实际: {real_w}x{real_h})"
            
            painter.setPen(QColor(255, 255, 255))
            painter.setFont(QFont("Arial", 12, QFont.Bold)) 
            text_y = self.selection_rect.y() - 10
            if text_y < 30: text_y = self.selection_rect.y() + 30
            painter.drawText(self.selection_rect.x(), text_y, info_text)
            
        else:
            painter.fillRect(self.rect(), bg_color)
            painter.setPen(QColor(255, 255, 255))
            painter.setFont(QFont("Arial", 16, QFont.Bold))
            hint = f"请框选区域 | 右键取消 | 缩放比: {self.scale_x:.2f}"
            fm = painter.fontMetrics()
            tw = fm.horizontalAdvance(hint)
            painter.drawText((self.width() - tw)//2, 100, hint)

        painter.setClipping(False)
        coord_text = f"Pos: {self.current_pos.x()},{self.current_pos.y()}"
        painter.setPen(QColor(255, 255, 0))
        painter.setFont(QFont("Arial", 10))
        painter.drawText(self.current_pos.x() + 20, self.current_pos.y() + 30, coord_text)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.start_point = event.pos()
            self.selection_rect = QRect()
            self.update()
        elif event.button() == Qt.RightButton:
            self.close()

    def mouseMoveEvent(self, event):
        self.current_pos = event.pos()
        if self.start_point:
            self.end_point = event.pos()
            self.selection_rect = QRect(self.start_point, self.end_point).normalized()
        self.update()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and self.start_point:
            rect = self.selection_rect
            self.close() 
            if rect.width() > 10 and rect.height() > 10:
                real_x = int(rect.x() * self.scale_x)
                real_y = int(rect.y() * self.scale_y)
                real_w = int(rect.width() * self.scale_x)
                real_h = int(rect.height() * self.scale_y)
                self.region_selected.emit((real_x, real_y, real_w, real_h))

# --------------------------
# 自定义帮助按钮
# --------------------------
class HelpBtn(QPushButton):
    def __init__(self, tip_text):
        super().__init__("?")
        self.setFixedSize(20, 20)
        self.setCursor(Qt.PointingHandCursor)
        self.setStyleSheet("""
            QPushButton {
                background-color: #2196F3; color: white; 
                border-radius: 10px; font-weight: bold; border: none;
            }
            QPushButton:hover { background-color: #1976D2; }
        """)
        self.tip_text = tip_text
        self.clicked.connect(self.show_tip)

    def show_tip(self):
        QToolTip.showText(QCursor.pos(), self.tip_text, self, QRect(), 5000)

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
                            if "任务管理器" in buff.value or "Task Manager" in buff.value:
                                self.trigger_stop("检测到【任务管理器】前台")
                                return
                time.sleep(0.02)
            except Exception as e:
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
                if task.get("type") not in [1.0, 2.0, 3.0, 8.0]: continue
                
                img = Image.open(path)
                img.load()
                self.img_cache[path] = img
                
                if self.min_scale != 1.0 or self.max_scale != 1.0:
                    if img.mode != 'L': img = img.convert('L')
                    template = np.array(img)
                    
                    templates_list = []
                    steps = int((self.max_scale - self.min_scale) / 0.05) + 1
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

        import cv2
        import numpy as np
        
        screen_np = np.array(screenshot_pil)
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

# --------------------------
# GUI 界面
# --------------------------
class WorkerThread(QThread):
    log_signal = Signal(str)
    finished_signal = Signal()
    def __init__(self, engine, tasks, loop_forever):
        super().__init__()
        self.engine = engine
        self.tasks = tasks
        self.loop_forever = loop_forever
        self.watchdog = None 

    def run(self):
        self.watchdog = FailsafeWatchdog(self.engine)
        self.watchdog.start()
        self.engine.run_tasks(self.tasks, self.loop_forever, self.log_callback)
        if self.watchdog: self.watchdog.kill()
        self.finished_signal.emit()

    def log_callback(self, msg): 
        if GLOBAL_CONFIG["log_to_ui"]:
            self.log_signal.emit(msg)

class TaskRow(QFrame):
    move_up_signal = Signal(object)
    move_down_signal = Signal(object)
    
    def __init__(self, delete_callback):
        super().__init__()
        self.parent_item = None
        self.setFrameShape(QFrame.StyledPanel)
        self.layout = QHBoxLayout(self)
        self.layout.setContentsMargins(2, 2, 2, 2)
        
        self.type_combo = QComboBox()
        self.type_combo.addItems(["左键单击", "左键双击", "右键单击", "输入文本", "等待(秒)", "滚轮滑动", "系统按键", "鼠标悬停", "截图保存"])
        self.type_combo.currentTextChanged.connect(self.on_type_changed)
        self.layout.addWidget(self.type_combo)
        
        self.value_input = QLineEdit()
        self.value_input.setPlaceholderText("参数")
        self.value_input.textChanged.connect(self.sync_data)
        self.layout.addWidget(self.value_input)
        
        self.file_btn = QPushButton("选择")
        self.file_btn.clicked.connect(self.select_file)
        self.layout.addWidget(self.file_btn)
        
        self.del_btn = QPushButton("X")
        self.del_btn.setStyleSheet("color: red; font-weight: bold;")
        self.del_btn.setFixedWidth(25)
        self.del_btn.clicked.connect(lambda: delete_callback(self))
        self.layout.addWidget(self.del_btn)
        
        self.on_type_changed(self.type_combo.currentText())

    def set_parent_item(self, item):
        self.parent_item = item
        self.sync_data() 

    def sync_data(self):
        if getattr(self, 'parent_item', None):
            self.parent_item.setData(Qt.UserRole, self.get_data())

    def on_type_changed(self, text):
        self.file_btn.setVisible("单击" in text or "悬停" in text or "截图" in text)
        self.sync_data()
            
    def set_data(self, data):
        self.value_input.setText(str(data.get("value", "")))
        TYPES_REV = {1.0: "左键单击", 2.0: "左键双击", 3.0: "右键单击", 4.0: "输入文本", 5.0: "等待(秒)", 6.0: "滚轮滑动", 7.0: "系统按键", 8.0: "鼠标悬停", 9.0: "截图保存"}
        t = data.get("type", 1.0)
        if t in TYPES_REV:
            self.type_combo.setCurrentText(TYPES_REV[t])

    def select_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择", filter="Images (*.png *.jpg *.bmp)")
        if path: self.value_input.setText(path)

    def get_data(self):
        TYPES = {"左键单击": 1.0, "左键双击": 2.0, "右键单击": 3.0, "输入文本": 4.0, "等待(秒)": 5.0, "滚轮滑动": 6.0, "系统按键": 7.0, "鼠标悬停": 8.0, "截图保存": 9.0}
        val = self.value_input.text()
        t = TYPES.get(self.type_combo.currentText(), 1.0)
        if t in [5.0, 6.0] and not val: val = "0"
        return {"type": t, "value": val}

class DraggableListWidget(QListWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDragDropMode(QAbstractItemView.InternalMove)
        self.setDefaultDropAction(Qt.MoveAction)
        self.setSelectionMode(QAbstractItemView.SingleSelection)

    def dropEvent(self, event):
        super().dropEvent(event)
        for i in range(self.count()):
            item = self.item(i)
            if self.itemWidget(item) is None:
                data = item.data(Qt.UserRole)
                if data:
                    self.window().restore_row_widget(item, data)

class RPAWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("不高兴就喝水 RPA配置工具(浮夸改V1.0)")
        self.resize(900, 850)
        self.engine = RPAEngine()
        self.settings = QSettings("MyRPA", "Config")
        self.hotkey_vk = 0x78 # 默认 F9
        
        self.current_process = None
        if HAS_PSUTIL:
            try: self.current_process = psutil.Process()
            except: pass
            
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        
        # 顶部
        top_bar = QHBoxLayout()
        add_btn = QPushButton("+ 新增指令")
        add_btn.clicked.connect(lambda: self.add_row())
        top_bar.addWidget(add_btn)
        save_btn = QPushButton("保存")
        save_btn.clicked.connect(self.save)
        top_bar.addWidget(save_btn)
        load_btn = QPushButton("导入")
        load_btn.clicked.connect(self.load)
        top_bar.addWidget(load_btn)
        
        # 设定区域
        region_btn = QPushButton("📷 设定识别区域")
        region_btn.setStyleSheet("background-color: #9C27B0; color: white; font-weight: bold;")
        region_btn.clicked.connect(self.open_region_selector)
        top_bar.addWidget(region_btn)
        
        top_bar.addStretch()
        main_layout.addLayout(top_bar)

        # 1. 识别配置
        g1 = QGroupBox("识别配置")
        gl1 = QHBoxLayout()
        gl1.addWidget(QLabel("相似:"))
        self.conf_edit = QLineEdit(self.settings.value("conf", "0.8")); self.conf_edit.setFixedWidth(50); gl1.addWidget(self.conf_edit)
        gl1.addWidget(HelpBtn("【相似度 (0.1 - 1.0)】\n数值越低：越容易匹配。\n数值越高：越精确。\nFlash游戏建议 0.6 - 0.8。"))
        gl1.addSpacing(20)
        gl1.addWidget(QLabel("缩放:"))
        self.scale_min = QLineEdit(self.settings.value("scale_min", "0.8")); self.scale_min.setFixedWidth(50); gl1.addWidget(self.scale_min)
        gl1.addWidget(QLabel("-")); 
        self.scale_max = QLineEdit(self.settings.value("scale_max", "1.2")); self.scale_max.setFixedWidth(50); gl1.addWidget(self.scale_max)
        gl1.addWidget(HelpBtn("【缩放范围】\n程序启动时会预先生成缩放模板缓存。\n范围越小，启动越快，内存占用越小。"))
        gl1.addStretch()
        g1.setLayout(gl1)
        main_layout.addWidget(g1)
        
        # 2. 避让设置
        g_dodge = QGroupBox("避让设置")
        gl_dodge = QHBoxLayout()
        gl_dodge.addWidget(QLabel("坐标1 X:"))
        self.dodge_x1 = QLineEdit(self.settings.value("dodge_x1", "100")); self.dodge_x1.setFixedWidth(50); gl_dodge.addWidget(self.dodge_x1)
        gl_dodge.addWidget(QLabel("Y:"))
        self.dodge_y1 = QLineEdit(self.settings.value("dodge_y1", "100")); self.dodge_y1.setFixedWidth(50); gl_dodge.addWidget(self.dodge_y1)
        gl_dodge.addSpacing(15)
        gl_dodge.addWidget(QLabel("坐标2 X:"))
        self.dodge_x2 = QLineEdit(self.settings.value("dodge_x2", "200")); self.dodge_x2.setFixedWidth(50); gl_dodge.addWidget(self.dodge_x2)
        gl_dodge.addWidget(QLabel("Y:"))
        self.dodge_y2 = QLineEdit(self.settings.value("dodge_y2", "100")); self.dodge_y2.setFixedWidth(50); gl_dodge.addWidget(self.dodge_y2)
        self.dodge_chk = QCheckBox("启用"); self.dodge_chk.setChecked(self.settings.value("dodge_en", False, type=bool))
        gl_dodge.addWidget(self.dodge_chk)
        self.double_dodge_chk = QCheckBox("二段"); self.double_dodge_chk.setChecked(self.settings.value("dbl_dodge", False, type=bool))
        gl_dodge.addWidget(self.double_dodge_chk)
        gl_dodge.addWidget(QLabel("间隔:"))
        self.dbl_wait = QLineEdit(self.settings.value("dbl_wait", "0.015")); self.dbl_wait.setFixedWidth(60); gl_dodge.addWidget(self.dbl_wait)
        gl_dodge.addWidget(HelpBtn("【二段避让】\n强迫游戏更新鼠标位置。"))
        gl_dodge.addStretch()
        g_dodge.setLayout(gl_dodge)
        main_layout.addWidget(g_dodge)
        
        # 3. 速度控制
        g2 = QGroupBox("速度控制 (0为极速)")
        gl2 = QHBoxLayout()
        gl2.addWidget(QLabel("移动(s):")); self.move_spd = QLineEdit(self.settings.value("move_spd", "0.0")); self.move_spd.setFixedWidth(50); gl2.addWidget(self.move_spd)
        gl2.addWidget(HelpBtn("【移动耗时】\n0.0=瞬移。"))
        gl2.addWidget(QLabel("按住(s):")); self.click_hld = QLineEdit(self.settings.value("click_hld", "0.04")); self.click_hld.setFixedWidth(50); gl2.addWidget(self.click_hld)
        gl2.addWidget(HelpBtn("【按住时长】\nFlash游戏建议 0.04-0.08。"))
        gl2.addWidget(QLabel("缓冲(s):")); self.settle = QLineEdit(self.settings.value("settle", "0.0")); self.settle.setFixedWidth(50); gl2.addWidget(self.settle)
        gl2.addWidget(HelpBtn("【结算缓冲】\n点击后的等待时间。"))
        gl2.addWidget(QLabel("超时(s):")); self.timeout = QLineEdit(self.settings.value("timeout", "0.0")); self.timeout.setFixedWidth(50); gl2.addWidget(self.timeout)
        gl2.addWidget(HelpBtn("【单步超时】\n0.0=扫一眼没找到直接过。"))
        gl2.addStretch()
        g2.setLayout(gl2)
        main_layout.addWidget(g2)
        
        # 4. 系统设置
        g3 = QGroupBox("系统设置")
        gl3 = QHBoxLayout()
        
        # 热键选择
        gl3.addWidget(QLabel("热键:"))
        self.hotkey_combo = QComboBox()
        self.hotkey_combo.addItems([f"F{i}" for i in range(1, 13)])
        saved_key = self.settings.value("hotkey", "F9")
        self.hotkey_combo.setCurrentText(saved_key)
        self.hotkey_combo.currentTextChanged.connect(self.update_hotkey_display)
        self.hotkey_combo.setFixedWidth(80)
        gl3.addWidget(self.hotkey_combo)
        
        self.tm_failsafe = QCheckBox("任务管理器急停"); self.tm_failsafe.setChecked(True); gl3.addWidget(self.tm_failsafe)
        self.tr_failsafe = QCheckBox("右上角急停"); self.tr_failsafe.setChecked(True); gl3.addWidget(self.tr_failsafe)
        self.key_failsafe = QCheckBox("ESC/中键急停"); self.key_failsafe.setChecked(True); gl3.addWidget(self.key_failsafe)
        
        gl3.addSpacing(15)
        self.log_file_chk = QCheckBox("写入文件日志"); 
        self.log_file_chk.setChecked(self.settings.value("log_file", False, type=bool))
        gl3.addWidget(self.log_file_chk)
        self.log_ui_chk = QCheckBox("显示界面日志"); 
        self.log_ui_chk.setChecked(self.settings.value("log_ui", True, type=bool))
        gl3.addWidget(self.log_ui_chk)
        self.log_file_chk.stateChanged.connect(self.update_log_config)
        self.log_ui_chk.stateChanged.connect(self.update_log_config)
        gl3.addStretch()
        g3.setLayout(gl3)
        main_layout.addWidget(g3)

        # 任务列表
        self.task_list = DraggableListWidget()
        main_layout.addWidget(self.task_list)
        
        # 底部
        bot_layout = QHBoxLayout()
        self.loop_combo = QComboBox(); self.loop_combo.addItems(["单次", "无限"])
        bot_layout.addWidget(self.loop_combo)
        self.mini_chk = QCheckBox("最小化"); 
        self.mini_chk.setChecked(self.settings.value("mini", False, type=bool))
        bot_layout.addWidget(self.mini_chk)
        
        self.start_btn = QPushButton("启动"); self.start_btn.clicked.connect(self.start_task)
        self.start_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        bot_layout.addWidget(self.start_btn)
        
        self.stop_btn = QPushButton("停止"); self.stop_btn.clicked.connect(self.stop_task)
        self.stop_btn.setStyleSheet("background-color: #f44336; color: white; font-weight: bold;")
        self.stop_btn.setEnabled(False)
        bot_layout.addWidget(self.stop_btn)
        
        main_layout.addLayout(bot_layout)
        
        self.log_text = QTextEdit(); self.log_text.setMaximumHeight(80)
        main_layout.addWidget(self.log_text)
        
        # 状态栏
        self.status_layout = QHBoxLayout()
        self.log_path_label = QLabel(f"日志: {get_log_path()}")
        self.log_path_label.setStyleSheet("color: gray; font-size: 10px;")
        main_layout.addWidget(self.log_path_label)
        
        self.status_layout = QHBoxLayout()
        self.region_label = QLabel("范围: 全屏")
        self.region_label.setStyleSheet("color: green;")
        self.status_layout.addWidget(self.region_label)
        self.status_layout.addStretch()
        self.cpu_label = QLabel("CPU: --")
        self.cpu_label.setStyleSheet("color: blue; font-weight: bold;")
        self.status_layout.addWidget(self.cpu_label)
        main_layout.addLayout(self.status_layout)
        
        self.add_row()
        self.cpu_timer = QTimer()
        self.cpu_timer.timeout.connect(self.update_cpu_info)
        self.cpu_timer.start(1000)
        self.update_log_config()
        self.update_hotkey_display(self.hotkey_combo.currentText())

        # 快捷键轮询
        self.hotkey_timer = QTimer()
        self.hotkey_timer.timeout.connect(self.check_hotkey)
        self.hotkey_timer.start(100)

    def update_hotkey_display(self, text):
        try:
            f_num = int(text.replace("F", ""))
            self.hotkey_vk = 0x70 + (f_num - 1)
            self.start_btn.setText(f"启动 ({text})")
            self.stop_btn.setText(f"停止 ({text})")
        except: pass

    def check_hotkey(self):
        if GetAsyncKeyState(self.hotkey_vk) & 0x8000:
            if self.engine.is_running:
                self.stop_task()
            else:
                self.start_task()
            self.hotkey_timer.stop()
            QTimer.singleShot(500, lambda: self.hotkey_timer.start(100))

    def open_region_selector(self):
        self.region_win = RegionWindow()
        self.region_win.region_selected.connect(self.on_region_selected)

    def on_region_selected(self, rect_tuple):
        self.engine.scan_region = rect_tuple
        self.region_label.setText(f"范围(物理): {rect_tuple}")
        self.log_text.append(f"已锁定游戏区域(物理): {rect_tuple} (速度+++)")

    def closeEvent(self, event):
        self.settings.setValue("conf", self.conf_edit.text())
        self.settings.setValue("scale_min", self.scale_min.text())
        self.settings.setValue("scale_max", self.scale_max.text())
        self.settings.setValue("dodge_x1", self.dodge_x1.text())
        self.settings.setValue("dodge_y1", self.dodge_y1.text())
        self.settings.setValue("dodge_x2", self.dodge_x2.text())
        self.settings.setValue("dodge_y2", self.dodge_y2.text())
        self.settings.setValue("dodge_en", self.dodge_chk.isChecked())
        self.settings.setValue("dbl_dodge", self.double_dodge_chk.isChecked())
        self.settings.setValue("dbl_wait", self.dbl_wait.text())
        self.settings.setValue("move_spd", self.move_spd.text())
        self.settings.setValue("click_hld", self.click_hld.text())
        self.settings.setValue("settle", self.settle.text())
        self.settings.setValue("timeout", self.timeout.text())
        self.settings.setValue("log_file", self.log_file_chk.isChecked())
        self.settings.setValue("log_ui", self.log_ui_chk.isChecked())
        self.settings.setValue("mini", self.mini_chk.isChecked())
        self.settings.setValue("hotkey", self.hotkey_combo.currentText())
        event.accept()

    def update_log_config(self):
        GLOBAL_CONFIG["log_to_file"] = self.log_file_chk.isChecked()
        GLOBAL_CONFIG["log_to_ui"] = self.log_ui_chk.isChecked()

    def update_cpu_info(self):
        core_str = "?"
        if HAS_KERNEL_CPU:
            try: core_str = str(GetCurrentProcessorNumber())
            except: pass
        sys_usage = "--"
        proc_usage = "--"
        if HAS_PSUTIL and self.current_process:
            try:
                sys_usage = f"{psutil.cpu_percent(interval=None):.1f}"
                raw_usage = self.current_process.cpu_percent(interval=None)
                proc_usage = f"{raw_usage:.1f}" 
            except: pass
        self.cpu_label.setText(f"逻辑核心: #{core_str} | 系统总占: {sys_usage}% | 脚本单核占: {proc_usage}%")

    def add_row(self, data=None):
        row_widget = TaskRow(delete_callback=self.del_row)
        if data: row_widget.set_data(data)
        item = QListWidgetItem(self.task_list)
        item.setSizeHint(row_widget.sizeHint())
        self.task_list.setItemWidget(item, row_widget)
        row_widget.set_parent_item(item)
        item.setData(Qt.UserRole, row_widget.get_data())

    def restore_row_widget(self, item, data):
        row_widget = TaskRow(delete_callback=self.del_row)
        row_widget.set_data(data)
        item.setSizeHint(row_widget.sizeHint())
        self.task_list.setItemWidget(item, row_widget)
        row_widget.set_parent_item(item)

    def del_row(self, row_widget):
        for i in range(self.task_list.count()):
            item = self.task_list.item(i)
            if self.task_list.itemWidget(item) == row_widget:
                self.task_list.takeItem(i)
                break

    def save(self):
        tasks = []
        for i in range(self.task_list.count()):
            item = self.task_list.item(i)
            widget = self.task_list.itemWidget(item)
            if widget: tasks.append(widget.get_data())
            else: tasks.append(item.data(Qt.UserRole))
        path, _ = QFileDialog.getSaveFileName(self, "保存", filter="JSON (*.json)")
        if path:
            with open(path, 'w', encoding='utf-8') as f: json.dump(tasks, f, ensure_ascii=False, indent=2)

    def load(self):
        path, _ = QFileDialog.getOpenFileName(self, "导入", filter="JSON (*.json)")
        if path:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                self.task_list.clear()
                for d in data: self.add_row(d)

    def start_task(self):
        tasks = []
        for i in range(self.task_list.count()):
            item = self.task_list.item(i)
            widget = self.task_list.itemWidget(item)
            if widget: tasks.append(widget.get_data())
        if not tasks: return
        try:
            self.engine.min_scale = float(self.scale_min.text())
            self.engine.max_scale = float(self.scale_max.text())
            self.engine.dodge_x1 = int(self.dodge_x1.text())
            self.engine.dodge_y1 = int(self.dodge_y1.text())
            self.engine.dodge_x2 = int(self.dodge_x2.text())
            self.engine.dodge_y2 = int(self.dodge_y2.text())
            self.engine.move_duration = float(self.move_spd.text())
            self.engine.click_hold = float(self.click_hld.text())
            self.engine.settlement_wait = float(self.settle.text())
            self.engine.timeout_val = float(self.timeout.text())
            self.engine.confidence = float(self.conf_edit.text())
            
            self.engine.enable_dodge = self.dodge_chk.isChecked()
            self.engine.enable_double_dodge = self.double_dodge_chk.isChecked()
            self.engine.double_dodge_wait = float(self.dbl_wait.text())
            
            self.engine.enable_tm_stop = self.tm_failsafe.isChecked()
            self.engine.enable_tr_stop = self.tr_failsafe.isChecked()
            self.engine.enable_key_stop = self.key_failsafe.isChecked()
        except: return QMessageBox.warning(self, "错误", "数值格式错误")

        if GLOBAL_CONFIG["log_to_ui"]:
            self.log_text.clear()
            self.log_text.append(f">>> 引擎启动({self.hotkey_combo.currentText()})...")
            
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        if self.mini_chk.isChecked(): self.showMinimized()
        
        is_loop = self.loop_combo.currentText() == "无限"
        self.worker = WorkerThread(self.engine, tasks, is_loop)
        self.worker.log_signal.connect(self.log_text.append)
        self.worker.finished_signal.connect(self.on_finish)
        self.worker.start()

    def stop_task(self):
        self.engine.stop()
        
    def on_finish(self):
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.showNormal()
        self.activateWindow()
        if GLOBAL_CONFIG["log_to_ui"]:
            self.log_text.append("结束")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = RPAWindow()
    win.show()
    sys.exit(app.exec())
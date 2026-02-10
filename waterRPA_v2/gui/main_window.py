# -*- coding: utf-8 -*-
import sys
import json
import ctypes
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                               QPushButton, QLabel, QComboBox, QLineEdit, 
                               QFileDialog, QMessageBox, QCheckBox, QGroupBox,
                               QTextEdit, QListWidgetItem)
from PySide6.QtCore import Qt, QTimer, QSettings

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

from ..config import GLOBAL_CONFIG
from ..utils import get_log_path
from ..engine import RPAEngine
from .widgets import RegionWindow, HelpBtn, TaskRow, DraggableListWidget, WorkerThread

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

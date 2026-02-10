# -*- coding: utf-8 -*-
from PySide6.QtWidgets import (QApplication, QWidget, QPushButton, QLabel, QComboBox, QLineEdit, 
                               QFileDialog, QFrame,  QToolTip, QListWidget, QListWidgetItem, QAbstractItemView, QHBoxLayout)
from PySide6.QtCore import Qt, Signal, QPoint, QRect, QThread, QSize
from PySide6.QtGui import QCursor, QFont, QColor, QPen, QPainter, QRegion
import pyautogui

from ..config import GLOBAL_CONFIG
from ..engine import FailsafeWatchdog

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

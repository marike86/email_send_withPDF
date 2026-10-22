"""基于 Qt 的 WinUI 3 风格界面；通过信号将后台结果交回主线程。"""

import argparse
import copy
from pathlib import Path
import sys
import threading

from PySide6.QtCore import QEvent, QPointF, Qt, QThread, Signal
from PySide6.QtGui import QColor, QFont, QPainter, QPen, QPolygonF
from PySide6.QtWidgets import (
    QApplication, QAbstractItemView, QComboBox, QDoubleSpinBox, QFileDialog,
    QFormLayout, QFrame, QHBoxLayout, QHeaderView, QLabel, QLineEdit,
    QListWidget, QMainWindow, QMessageBox, QProgressBar, QPushButton,
    QScrollArea, QSpinBox, QSplitter, QStackedWidget, QTableWidget,
    QTableWidgetItem, QTextEdit, QVBoxLayout, QWidget, QStyle, QStyleOptionSpinBox,
    QStyledItemDelegate,
)

import send_email as core

STYLE = """
QWidget { font-family: 'Segoe UI', 'Microsoft YaHei UI'; font-size: 13px; color: #202020; }
QMainWindow, #content { background: #f7f8fa; }
#sidebar { background: #f0f2f5; border-right: 1px solid #e2e5e9; }
#brand { font-size: 20px; font-weight: 600; color: #005fb8; }
#title { font-size: 28px; font-weight: 600; }
#subtitle, #hint { color: #657081; }
#card { background: white; border: 1px solid #e4e7eb; border-radius: 10px; }
#section { font-size: 16px; font-weight: 600; }
#metric { font-size: 26px; font-weight: 600; color: #005fb8; }
QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox {
    background: white; border: 1px solid #d7dce2; border-bottom: 1px solid #87919e;
    border-radius: 5px; padding: 7px 9px; min-height: 18px;
}
QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus { border-bottom: 2px solid #0067c0; }
QComboBox { padding-right: 36px; }
QComboBox:hover, QSpinBox:hover, QDoubleSpinBox:hover { border-color: #b2bac5; border-bottom-color: #677587; }
QComboBox:focus, QComboBox:on { border-bottom: 2px solid #0067c0; }
QComboBox::drop-down { subcontrol-origin: border; subcontrol-position: top right; width: 34px; border: none; background: transparent; }
QComboBox::down-arrow { image: none; }
QComboBox QAbstractItemView { background: white; border: 1px solid #d7dce2; border-radius: 6px; padding: 4px; outline: none; selection-background-color: #e6f2ff; selection-color: #005fb8; }
QComboBox QAbstractItemView::item { min-height: 30px; padding: 3px 10px; border-radius: 4px; }
QSpinBox, QDoubleSpinBox { padding-right: 36px; }
QSpinBox QLineEdit, QDoubleSpinBox QLineEdit { border: none; border-radius: 0; background: transparent; padding: 0; min-height: 0; }
QSpinBox::up-button, QDoubleSpinBox::up-button { subcontrol-origin: border; subcontrol-position: top right; width: 30px; height: 17px; border: none; background: transparent; margin: 1px 1px 0 0; }
QSpinBox::down-button, QDoubleSpinBox::down-button { subcontrol-origin: border; subcontrol-position: bottom right; width: 30px; height: 17px; border: none; background: transparent; margin: 0 1px 1px 0; }
QSpinBox::up-button:hover, QSpinBox::down-button:hover, QDoubleSpinBox::up-button:hover, QDoubleSpinBox::down-button:hover { background: #edf3fa; border-radius: 4px; }
QSpinBox::up-arrow, QSpinBox::down-arrow, QDoubleSpinBox::up-arrow, QDoubleSpinBox::down-arrow { image: none; }
QFrame#pathField { background: white; border: 1px solid #d7dce2; border-bottom: 1px solid #87919e; border-radius: 5px; }
QFrame#pathField:hover { border-color: #b2bac5; border-bottom-color: #677587; }
QFrame#pathField[focused="true"] { border-bottom: 2px solid #0067c0; }
QLineEdit#pathInput, QLineEdit#pathInput:focus { background: transparent; border: none; padding: 4px 9px; min-height: 0; }
QPushButton#pathBrowse { background: #f8fafc; border: none; border-left: 1px solid #e4e7eb; border-radius: 3px; padding: 4px 12px; }
QPushButton#pathBrowse:hover { background: #eaf2fb; }
QPushButton { background: white; border: 1px solid #d7dce2; border-radius: 5px; padding: 8px 15px; }
QPushButton:hover { background: #f0f6fc; border-color: #9bbddd; }
QPushButton:pressed { background: #e2effb; }
QPushButton:disabled { color: #9a9fa6; background: #f4f5f7; border-color: #e6e8eb; }
QPushButton#primary { background: #0067c0; color: white; border-color: #0067c0; }
QPushButton#primary:hover { background: #005aa9; }
QPushButton#primary:disabled { background: #b2cce3; border-color: #b2cce3; }
QListWidget { background: transparent; border: none; outline: none; }
QListWidget::item { padding: 13px 18px; margin: 3px 10px; border-radius: 6px; }
QListWidget::item:selected { background: #e1eaf3; color: #005fb8; font-weight: 600; }
QListWidget::item:hover { background: #e6ebf1; }
QTableWidget { background: white; alternate-background-color: #fafbfd; border: none; gridline-color: #edf0f4; }
QTableWidget::item { padding: 8px; border-bottom: 1px solid #edf0f4; }
QTableWidget::item:selected { background: #e6f2ff; color: #202020; }
QHeaderView::section { background: #f6f8fa; border: none; border-bottom: 1px solid #e4e7eb; padding: 10px; color: #586474; }
QTextEdit { background: white; border: none; padding: 8px; }
QProgressBar { background: #e4eaf0; border: none; border-radius: 3px; max-height: 6px; }
QProgressBar::chunk { background: #0067c0; border-radius: 3px; }
QScrollArea { border: none; background: transparent; }
QSplitter::handle { background: #f7f8fa; width: 12px; }
"""


def paint_chevron(painter, center, up=False, enabled=True):
    """以矢量线条绘制清晰的箭头，不依赖系统字体或额外图片。"""
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    painter.setPen(QPen(QColor("#526174" if enabled else "#aab2bc"), 1.4,
                        Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
    direction = -1 if up else 1
    x, y = center.x(), center.y()
    painter.drawPolyline(QPolygonF([QPointF(x - 3.5, y - 1.5 * direction),
                                   QPointF(x, y + 2 * direction),
                                   QPointF(x + 3.5, y - 1.5 * direction)]))


class FluentComboBox(QComboBox):
    """统一下拉箭头与列表项间距，避免 Qt 原生子控件边线混入。"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setItemDelegate(QStyledItemDelegate(self))
        self.setMinimumHeight(36)

    def paintEvent(self, event):
        super().paintEvent(event)
        painter = QPainter(self)
        paint_chevron(painter, QPointF(self.width() - 18, self.height() / 2), enabled=self.isEnabled())
        painter.end()


class FluentSpinMixin:
    """保留数值增减按钮行为，仅统一其箭头绘制。"""
    def paintEvent(self, event):
        super().paintEvent(event)
        option = QStyleOptionSpinBox()
        self.initStyleOption(option)
        painter = QPainter(self)
        for control, up in ((QStyle.SubControl.SC_SpinBoxUp, True), (QStyle.SubControl.SC_SpinBoxDown, False)):
            rect = self.style().subControlRect(QStyle.ComplexControl.CC_SpinBox, option, control, self)
            paint_chevron(painter, QPointF(rect.center()), up, self.isEnabled())
        painter.end()


class FluentSpinBox(FluentSpinMixin, QSpinBox):
    """统一风格的整数输入框。"""


class FluentDoubleSpinBox(FluentSpinMixin, QDoubleSpinBox):
    """统一风格的小数输入框。"""


class PathField(QFrame):
    """一体式路径输入与浏览按钮，悬停显示完整路径。"""
    def __init__(self, value, title, browse, parent=None):
        super().__init__(parent)
        self.setObjectName("pathField")
        self.setFixedHeight(38)
        row = QHBoxLayout(self)
        row.setContentsMargins(2, 2, 2, 2)
        row.setSpacing(0)
        self.edit = QLineEdit(value)
        self.edit.setObjectName("pathInput")
        self.edit.setPlaceholderText("请选择" + title)
        self.edit.setToolTip(value)
        self.edit.setCursorPosition(0)
        self.edit.textChanged.connect(self.edit.setToolTip)
        self.edit.editingFinished.connect(lambda: self.edit.setCursorPosition(0))
        self.edit.installEventFilter(self)
        row.addWidget(self.edit, 1)
        self.browse_button = QPushButton("浏览…")
        self.browse_button.setObjectName("pathBrowse")
        self.browse_button.setToolTip("选择" + title)
        self.browse_button.clicked.connect(browse)
        row.addWidget(self.browse_button)
        self.setFocusProxy(self.edit)

    def eventFilter(self, watched, event):
        if watched is self.edit and event.type() in (QEvent.Type.FocusIn, QEvent.Type.FocusOut):
            self.setProperty("focused", event.type() == QEvent.Type.FocusIn)
            self.style().unpolish(self)
            self.style().polish(self)
        return super().eventFilter(watched, event)


def label(text, name=None):
    """创建可换行的界面标签。"""
    widget = QLabel(text)
    widget.setWordWrap(True)
    if name:
        widget.setObjectName(name)
    return widget


def card():
    """创建 Fluent 风格白色卡片及其布局。"""
    frame = QFrame()
    frame.setObjectName("card")
    layout = QVBoxLayout(frame)
    layout.setContentsMargins(18, 16, 18, 16)
    layout.setSpacing(12)
    return frame, layout


class Worker(QThread):
    """后台执行预检或发送；不从工作线程读取或修改控件。"""
    result = Signal(object)
    updated = Signal(object, int, int)
    error = Signal(str)

    def __init__(self, action, parent):
        super().__init__(parent)
        self.action = action

    def run(self):
        try:
            self.result.emit(self.action(self.updated.emit))
        except Exception as exc:
            self.error.emit(str(exc))


class MainWindow(QMainWindow):
    """邮件工作台：文件预检、逐封预览、断点续发和设置。"""

    def __init__(self, config_path="config.json", state_path=core.STATUS_FILE):
        super().__init__()
        self.config_path = str(Path(config_path).resolve())
        self.state_path = str(Path(state_path).resolve())
        self.config = core.load_config(self.config_path)
        self.jobs = []
        self.snapshot = None
        self.worker = None
        self.stop = threading.Event()
        self.setWindowTitle("PDF 邮件工作台")
        self.resize(1240, 900)
        self.setMinimumSize(1000, 760)
        root = QWidget()
        outer = QHBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)
        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(190)
        side = QVBoxLayout(sidebar)
        side.setContentsMargins(0, 28, 0, 20)
        brand = label("   PDF 邮件", "brand")
        side.addWidget(brand)
        side.addSpacing(26)
        self.nav = QListWidget()
        self.nav.addItems(["发送工作台", "账户与设置"])
        side.addWidget(self.nav)
        side.addWidget(label("    PDF MAILER\n    准确匹配 · 从容发送", "hint"))
        outer.addWidget(sidebar)
        self.pages = QStackedWidget()
        self.pages.setObjectName("content")
        outer.addWidget(self.pages, 1)
        self.fields = {}
        self.build_workspace()
        self.build_settings()
        self.nav.currentRowChanged.connect(self.pages.setCurrentIndex)
        self.nav.setCurrentRow(0)
        self.setCentralWidget(root)
        for field in self.fields.values():
            if isinstance(field, QLineEdit):
                field.textChanged.connect(self.invalidate)
            elif isinstance(field, QComboBox):
                field.currentIndexChanged.connect(self.invalidate)
            else:
                field.valueChanged.connect(self.invalidate)
        self.update_actions()

    def page(self, title, subtitle):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(14)
        layout.addWidget(label(title, "title"))
        layout.addWidget(label(subtitle, "subtitle"))
        self.pages.addWidget(widget)
        return layout

    def path_field(self, layout, title, key, kind):
        row = QHBoxLayout()
        text = label(title)
        text.setFixedWidth(68)
        row.addWidget(text)
        field = PathField(str(self.config.get(key, "")), title, lambda: self.browse(key, kind))
        self.fields[key] = field.edit
        row.addWidget(field, 1)
        layout.addLayout(row)

    def build_workspace(self):
        layout = self.page("发送工作台", "将每一份 PDF，送到对应的人手中。先检查，再确认发送。")
        frame, body = card()
        body.addWidget(label("01   准备文件", "section"))
        self.path_field(body, "PDF 目录", "attachment_folder", "folder")
        self.path_field(body, "学生表格", "student_info_excel", "excel")
        self.path_field(body, "邮件模板", "email_template", "template")
        self.files_card = frame
        layout.addWidget(frame)
        metrics = QHBoxLayout()
        self.metrics = []
        for title in ("附件总数", "待发送", "已发送", "需关注"):
            frame, body = card()
            number = label("0", "metric")
            body.addWidget(number)
            body.addWidget(label(title, "hint"))
            self.metrics.append(number)
            metrics.addWidget(frame)
        layout.addLayout(metrics)
        toolbar = QHBoxLayout()
        toolbar.addWidget(label("02   检查与预览", "section"))
        toolbar.addStretch()
        self.search = QLineEdit()
        self.search.setPlaceholderText("搜索姓名、学号或文件…")
        self.search.setMaximumWidth(240)
        self.search.textChanged.connect(self.filter_rows)
        toolbar.addWidget(self.search)
        self.filter = FluentComboBox()
        self.filter.addItems(["全部状态", "待发送", "已发送", "需关注"])
        self.filter.currentIndexChanged.connect(self.filter_rows)
        toolbar.addWidget(self.filter)
        layout.addLayout(toolbar)
        split = QSplitter()
        table_frame, table_body = card()
        table_body.setContentsMargins(0, 0, 0, 0)
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["PDF 文件", "姓名", "收件邮箱", "状态", "说明"])
        self.table.verticalHeader().hide()
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setShowGrid(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setColumnWidth(0, 180)
        self.table.setColumnWidth(1, 75)
        self.table.setColumnWidth(2, 200)
        self.table.setColumnWidth(3, 80)
        self.table.itemSelectionChanged.connect(self.show_preview)
        table_body.addWidget(self.table)
        split.addWidget(table_frame)
        preview_frame, preview_body = card()
        preview_body.addWidget(label("邮件预览", "section"))
        self.preview = QTextEdit()
        self.preview.setReadOnly(True)
        self.preview.setPlaceholderText("完成预检后，选择一行查看收件人、主题和正文。\n\n不会在预检时连接邮箱。")
        preview_body.addWidget(self.preview)
        self.resolve = QPushButton("处理待核实记录…")
        self.resolve.clicked.connect(self.resolve_uncertain)
        self.resolve.setEnabled(False)
        preview_body.addWidget(self.resolve)
        split.addWidget(preview_frame)
        split.setSizes([660, 280])
        layout.addWidget(split, 1)
        self.notice = label("请选择文件，然后点击“检查文件”。授权码在账户与设置中填写。", "hint")
        layout.addWidget(self.notice)
        self.progress = QProgressBar()
        self.progress.setTextVisible(False)
        self.progress.setValue(0)
        layout.addWidget(self.progress)
        actions = QHBoxLayout()
        self.check = QPushButton("检查文件")
        self.check.clicked.connect(self.preflight)
        self.export = QPushButton("导出结果")
        self.export.clicked.connect(self.export_csv)
        self.retry = QPushButton("重试失败项")
        self.retry.clicked.connect(lambda: self.send(True))
        self.cancel = QPushButton("停止发送")
        self.cancel.clicked.connect(self.request_stop)
        self.start = QPushButton("开始发送")
        self.start.setObjectName("primary")
        self.start.clicked.connect(lambda: self.send(False))
        for button in (self.check, self.export, self.retry):
            actions.addWidget(button)
        actions.addStretch()
        actions.addWidget(self.cancel)
        actions.addWidget(self.start)
        layout.addLayout(actions)

    def build_settings(self):
        layout = self.page("账户与设置", "授权码仅用于本次会话。保存设置时不会写入授权码。")
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content = QWidget()
        body = QVBoxLayout(content)
        body.setContentsMargins(0, 0, 0, 0)
        for title, items in [
            ("发件账户", [("sender_email", "发件邮箱", "text"), ("sender_password", "邮箱授权码", "password"),
                         ("smtp_server", "SMTP 服务器", "text"), ("smtp_port", "端口", "port"), ("security", "加密方式", "security")]),
            ("Excel 列映射", [("excel_student_id_col", "学号列", "column"), ("excel_name_col", "姓名列", "column"),
                              ("excel_email_col", "邮箱列", "column")]),
            ("发送节奏", [("send_interval", "邮件间隔（秒）", "seconds"), ("max_retries", "最多尝试次数", "retries"),
                          ("retry_interval", "重试间隔（秒）", "seconds")]),
        ]:
            frame, box = card()
            box.addWidget(label(title, "section"))
            form = QFormLayout()
            form.setSpacing(12)
            for key, caption, kind in items:
                value = self.config.get(key, "")
                if kind in {"text", "password", "column"}:
                    field = QLineEdit(str(value))
                    if kind == "password":
                        field.setEchoMode(QLineEdit.EchoMode.Password)
                        field.setPlaceholderText("邮箱授权码 / 应用专用密码")
                    if kind == "column":
                        field.setText(core.openpyxl.utils.get_column_letter(core.column_number(value)))
                elif kind == "security":
                    field = FluentComboBox()
                    field.addItems(["STARTTLS · 通常使用 587", "SSL / TLS · 通常使用 465"])
                    field.setCurrentIndex(1 if self.config.get("use_ssl") or int(self.config["smtp_port"]) == 465 else 0)
                elif kind == "seconds":
                    field = FluentDoubleSpinBox()
                    field.setRange(0, 86400)
                    field.setValue(float(value))
                else:
                    field = FluentSpinBox()
                    field.setRange(1, 65535 if kind == "port" else 10)
                    field.setValue(int(value))
                field.setFixedHeight(36)
                self.fields[key] = field
                form.addRow(caption, field)
            box.addLayout(form)
            if title == "Excel 列映射":
                box.addWidget(label("第一行为表头；列名支持 A、AA 等。含前导零的学号请在 Excel 中存为文本。", "hint"))
            body.addWidget(frame)
        state = label(f"发送日志：{self.state_path}\n配置文件：{self.config_path}", "hint")
        state.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        body.addWidget(state)
        body.addStretch()
        scroll.setWidget(content)
        layout.addWidget(scroll)
        self.settings_content = content
        save = QPushButton("保存设置")
        save.setObjectName("primary")
        save.clicked.connect(self.save_settings)
        layout.addWidget(save, alignment=Qt.AlignmentFlag.AlignRight)
        self.save_button = save

    def collect_config(self):
        config = dict(self.config)
        for key, field in self.fields.items():
            if isinstance(field, QLineEdit):
                config[key] = field.text().strip()
            elif isinstance(field, QComboBox):
                config["use_ssl"] = field.currentIndex() == 1
                config["use_tls"] = field.currentIndex() == 0
            else:
                config[key] = field.value()
        config.pop("security", None)
        for key in ("excel_student_id_col", "excel_name_col", "excel_email_col"):
            config[key] = core.column_number(config[key])
        core.validate_config(config)
        return config

    def browse(self, key, kind):
        if kind == "folder":
            path = QFileDialog.getExistingDirectory(self, "选择 PDF 附件目录")
        else:
            path, _ = QFileDialog.getOpenFileName(self, "选择文件", "", "Excel (*.xlsx)" if kind == "excel" else "文本模板 (*.txt)")
        if path:
            self.fields[key].setText(path)
            self.fields[key].setCursorPosition(0)

    def save_settings(self):
        try:
            self.config = self.collect_config()
            core.save_config(self.config, self.config_path)
            QMessageBox.information(self, "设置已保存", "已保存配置。授权码未写入文件。")
        except Exception as exc:
            self.show_error(str(exc))

    def invalidate(self, *_):
        self.snapshot = None
        self.notice.setText("设置已更改，请重新检查文件以更新发送计划。")
        self.update_actions()

    def busy(self):
        return self.worker is not None

    def update_actions(self):
        busy = self.busy()
        ready = self.snapshot is not None
        self.check.setEnabled(not busy)
        self.export.setEnabled(bool(self.jobs) and not busy)
        self.start.setEnabled(ready and not busy and any(j["status"] in {"ready", "failed"} for j in self.jobs))
        self.retry.setEnabled(ready and not busy and any(j["status"] == "failed" for j in self.jobs))
        self.cancel.setEnabled(busy and self.snapshot is not None)
        self.files_card.setEnabled(not busy)
        self.settings_content.setEnabled(not busy)
        self.save_button.setEnabled(not busy)
        self.show_preview()

    def launch(self, action, callback):
        self.worker = Worker(action, self)
        self.worker.result.connect(callback)
        self.worker.error.connect(self.show_error)
        self.worker.updated.connect(self.on_update)
        self.worker.finished.connect(self.worker_finished)
        self.update_actions()
        self.worker.start()

    def worker_finished(self):
        worker = self.worker
        self.worker = None
        worker.deleteLater()
        self.update_actions()

    def preflight(self):
        try:
            config = self.collect_config()
            self.snapshot = None
            self.notice.setText("正在检查附件、Excel 和发送日志…")
            self.progress.setRange(0, 0)
            self.launch(lambda _: core.prepare_batch(config, config["attachment_folder"], self.state_path),
                        lambda jobs: self.preflight_done(config, jobs))
        except Exception as exc:
            self.show_error(str(exc))

    def preflight_done(self, config, jobs):
        self.snapshot = config
        self.jobs = jobs
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.render_jobs()
        attention = sum(j["status"] in {"blocked", "failed", "uncertain", "sending"} for j in jobs)
        self.notice.setText(f"检查完成：{len(jobs)} 个 PDF，{attention} 项需关注。选择一行可预览邮件。" if jobs else "该目录未找到 PDF 文件。")

    def render_jobs(self):
        selected = self.table.currentRow()
        self.table.setRowCount(len(self.jobs))
        colors = {"success": "#107c41", "ready": "#005fb8", "blocked": "#b42318", "failed": "#b42318", "uncertain": "#925b00", "sending": "#925b00"}
        for row, job in enumerate(self.jobs):
            values = [job["filename"], job["name"], job["email"], core.LABELS[job["status"]], job["error"]]
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                item.setToolTip(str(value))
                if column == 3:
                    item.setForeground(QColor(colors[job["status"]]))
                self.table.setItem(row, column, item)
            self.table.setRowHeight(row, 42)
        counts = [len(self.jobs), sum(j["status"] in {"ready", "failed"} for j in self.jobs),
                  sum(j["status"] == "success" for j in self.jobs),
                  sum(j["status"] in {"blocked", "failed", "uncertain", "sending"} for j in self.jobs)]
        for widget, count in zip(self.metrics, counts):
            widget.setText(str(count))
        self.filter_rows()
        if self.jobs:
            self.table.selectRow(max(0, min(selected, len(self.jobs) - 1)))
        self.show_preview()

    def filter_rows(self, *_):
        needle = self.search.text().strip().lower()
        index = self.filter.currentIndex()
        states = [set(core.LABELS), {"ready", "failed"}, {"success"}, {"blocked", "failed", "uncertain", "sending"}][index]
        for row, job in enumerate(self.jobs):
            match = needle in " ".join(str(job.get(k, "")) for k in ("filename", "student_id", "name", "email")).lower()
            self.table.setRowHidden(row, not match or job["status"] not in states)

    def show_preview(self):
        row = self.table.currentRow()
        job = self.jobs[row] if 0 <= row < len(self.jobs) else None
        self.resolve.setEnabled(bool(job and job["status"] in {"uncertain", "sending"} and not self.busy() and self.snapshot))
        if job:
            self.preview.setPlainText(f"收件人：{job['name']} <{job['email']}>\n附件：{job['filename']}\n\n主题：{job['subject']}\n\n{job['body']}\n\n{job['error']}")
        else:
            self.preview.clear()

    def send(self, retry_only=False):
        try:
            if self.snapshot is None or self.busy():
                return
            config = dict(self.snapshot)
            core.validate_config(config, sending=True)
            states = {"failed"} if retry_only else {"ready", "failed"}
            count = sum(j["status"] in states for j in self.jobs)
            reply = QMessageBox.question(self, "确认发送计划",
                f"发件人：{config['sender_email']}\n将向匹配收件人发送 {count} 封带 PDF 的邮件。\n已发送、需修正和待核实的记录会跳过。\n\n请确认已核对预览内容。",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
            if reply != QMessageBox.StandardButton.Yes:
                return
            self.stop.clear()
            jobs = copy.deepcopy(self.jobs)
            self.notice.setText("正在发送…可停止后续邮件；当前网络请求完成后停止。")
            self.progress.setRange(0, count)
            self.progress.setValue(0)
            self.launch(lambda update: core.run_batch(config, jobs, self.state_path, self.stop, update, retry_only), self.send_done)
        except Exception as exc:
            self.show_error(str(exc))

    def on_update(self, job, index, total):
        for row, existing in enumerate(self.jobs):
            if existing["key"] == job["key"]:
                self.jobs[row] = dict(job)
        self.render_jobs()
        self.progress.setRange(0, total)
        self.progress.setValue(index)
        self.notice.setText(f"{index} / {total} · {core.LABELS[job['status']]} · {job['filename']}")

    def send_done(self, jobs):
        self.jobs = jobs
        self.render_jobs()
        self.notice.setText("任务已停止，进度已保存。可继续发送剩余项。" if self.stop.is_set() else "本轮处理完成。可导出结果，或修正问题后重新检查。")

    def request_stop(self):
        self.stop.set()
        self.cancel.setEnabled(False)
        self.notice.setText("正在停止后续发送，等待当前网络请求完成…")

    def resolve_uncertain(self):
        row = self.table.currentRow()
        if not 0 <= row < len(self.jobs):
            return
        job = self.jobs[row]
        box = QMessageBox(self)
        box.setWindowTitle("核实发送结果")
        box.setText(f"请先在邮箱中检查：{job['filename']} → {job['email']}\n确认结果后选择处理方式。")
        sent = box.addButton("已送达，标记成功", QMessageBox.ButtonRole.AcceptRole)
        retry = box.addButton("确认未送达，允许重试", QMessageBox.ButtonRole.DestructiveRole)
        box.addButton("取消", QMessageBox.ButtonRole.RejectRole)
        box.exec()
        if box.clickedButton() not in (sent, retry):
            return
        try:
            with core.status_lock(self.state_path):
                status = core.load_status(self.state_path)
                record = status["records"][job["key"]]
                if record["status"] not in {"uncertain", "sending"}:
                    raise ValueError("记录已经变化，请重新检查文件")
                record.update(status="success" if box.clickedButton() == sent else "failed", error="用户已核实发送结果")
                core.save_status(self.state_path, status)
            job.update(status=record["status"], error=record["error"])
            self.render_jobs()
            self.update_actions()
        except Exception as exc:
            self.show_error(str(exc))

    def export_csv(self):
        path, _ = QFileDialog.getSaveFileName(self, "导出全部检查与发送结果", "邮件发送结果.xlsx", "Excel (*.xlsx);;CSV (*.csv)")
        if path:
            try:
                core.export_results(self.jobs, path)
                self.notice.setText(f"结果已导出：{path}")
            except Exception as exc:
                self.show_error(str(exc))

    def show_error(self, message):
        self.progress.setRange(0, 100)
        self.notice.setText("操作未完成：" + message)
        # 发生错误后必须重新读取日志，避免内存状态与落盘结果不一致。
        self.snapshot = None
        self.update_actions()
        QMessageBox.warning(self, "请检查", message)

    def closeEvent(self, event):
        if self.busy():
            self.request_stop()
            self.notice.setText("请等待当前任务停止后再关闭窗口。")
            event.ignore()
        else:
            event.accept()


def main():
    """启动桌面界面，支持指定配置和状态文件。"""
    parser = argparse.ArgumentParser(description="PDF 邮件工作台")
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--state", default=core.STATUS_FILE)
    args = parser.parse_args()
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setFont(QFont("Segoe UI", 10))
    app.setStyleSheet(STYLE)
    try:
        window = MainWindow(args.config, args.state)
    except Exception as exc:
        QMessageBox.critical(None, "无法读取配置", str(exc))
        return 1
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())

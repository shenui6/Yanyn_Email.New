import sys
import os
import logging
import socket
import imaplib
import poplib
import email
import smtplib
import shutil
import re
import threading
import win32event
import winerror
import win32api
import qtawesome as qta
from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import decode_header
from email.utils import parseaddr
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QLineEdit, QPushButton, QTextEdit, QListWidget,
                             QTabWidget, QTreeWidget, QTreeWidgetItem, QSplitter,
                             QToolBar, QStatusBar, QDialog, QFormLayout, QMessageBox,
                             QFileDialog, QComboBox, QCheckBox, QGroupBox, QRadioButton,
                             QDialogButtonBox, QSystemTrayIcon, QMenu, QStyle)
from PyQt6.QtCore import Qt, QSettings, QStandardPaths, QTimer, QMutex, QMutexLocker
from PyQt6.QtGui import QIcon, QColor, QPalette, QTextCursor, QAction, QFont
from PyQt6.QtWebEngineWidgets import QWebEngineView

from UpdateEmailListSignal import UpdateEmailListSignal

# 设置日志
logging.basicConfig(
    filename='yanyn_email.log',
    level=logging.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# 防止多开 - Windows版本
mutex_name = "YanynEmailMutex"
mutex = win32event.CreateMutex(None, False, mutex_name)
last_error = win32api.GetLastError()

if last_error == winerror.ERROR_ALREADY_EXISTS:
    QMessageBox.critical(None, "错误", "Yanyn Email 已经在运行中!")
    sys.exit(1)


class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("设置")
        self.setWindowIcon(QIcon("icon.ico"))  # 添加设置图标
        self.resize(400, 300)

        # 使用垂直布局
        layout = QVBoxLayout(self)

        # 配色方案选择
        self.theme_group = QGroupBox("配色方案")
        theme_layout = QVBoxLayout()

        self.theme_azure = QRadioButton("蔚蓝色主题 (默认)")
        self.theme_light = QRadioButton("浅色主题")
        self.theme_dark = QRadioButton("深色主题")

        # 设置当前选中的主题
        current_theme = self.parent().settings.value("theme", "azure")
        if current_theme == "azure":
            self.theme_azure.setChecked(True)
        elif current_theme == "light":
            self.theme_light.setChecked(True)
        else:
            self.theme_dark.setChecked(True)

        theme_layout.addWidget(self.theme_azure)
        theme_layout.addWidget(self.theme_light)
        theme_layout.addWidget(self.theme_dark)
        self.theme_group.setLayout(theme_layout)

        # 托盘选项
        self.tray_group = QGroupBox("系统托盘")
        tray_layout = QVBoxLayout()

        self.minimize_to_tray = QCheckBox("最小化到托盘")
        self.minimize_to_tray.setChecked(self.parent().settings.value("minimize_to_tray", True, type=bool))

        tray_layout.addWidget(self.minimize_to_tray)
        self.tray_group.setLayout(tray_layout)

        # 按钮
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        # 添加到主布局
        layout.addWidget(self.theme_group)
        layout.addWidget(self.tray_group)
        layout.addWidget(self.button_box)

    def get_settings(self):
        """返回设置选项"""
        theme = "azure"
        if self.theme_light.isChecked():
            theme = "light"
        elif self.theme_dark.isChecked():
            theme = "dark"

        return {
            "theme": theme,
            "minimize_to_tray": self.minimize_to_tray.isChecked()
        }


class ComposeEmailDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("写邮件")
        self.setWindowIcon(QIcon("icon.ico"))
        self.resize(600, 500)

        layout = QVBoxLayout(self)

        # 发件人标签（只读）
        self.from_label = QLabel()
        layout.addWidget(QLabel("发件人:"))
        layout.addWidget(self.from_label)

        # 收件人
        self.to_edit = QLineEdit()
        layout.addWidget(QLabel("收件人:"))
        layout.addWidget(self.to_edit)

        # 主题
        self.subject_edit = QLineEdit()
        layout.addWidget(QLabel("主题:"))
        layout.addWidget(self.subject_edit)

        # 邮件内容
        self.content_edit = QTextEdit()
        layout.addWidget(QLabel("内容:"))
        layout.addWidget(self.content_edit)

        # 附件
        self.attachment_label = QLabel("无附件")
        self.attachment_btn = QPushButton("添加附件")
        self.attachment_btn.clicked.connect(self.add_attachment)

        attachment_layout = QHBoxLayout()
        attachment_layout.addWidget(QLabel("附件:"))
        attachment_layout.addWidget(self.attachment_label)
        attachment_layout.addWidget(self.attachment_btn)
        layout.addLayout(attachment_layout)

        # 按钮
        self.button_box = QHBoxLayout()
        self.send_btn = QPushButton("发送")
        self.send_btn.clicked.connect(self.accept)
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.clicked.connect(self.reject)
        self.button_box.addWidget(self.send_btn)
        self.button_box.addWidget(self.cancel_btn)
        layout.addLayout(self.button_box)

        self.attachment_path = None

    def set_from_email(self, email):
        """设置发件人邮箱"""
        self.from_label.setText(email)

    def add_attachment(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择附件", "", "所有文件 (*.*)")
        if path:
            self.attachment_path = path
            self.attachment_label.setText(os.path.basename(path))

    def get_email_data(self):
        return {
            "to": self.to_edit.text(),
            "subject": self.subject_edit.text(),
            "content": self.content_edit.toPlainText(),
            "attachment": self.attachment_path
        }


class AddAccountDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("添加邮箱账户")
        self.setWindowIcon(QIcon("icon.ico"))
        self.resize(500, 300)

        layout = QFormLayout(self)

        # 预设邮箱选择
        self.preset_combo = QComboBox()
        self.preset_combo.addItems(["晏阳邮箱", "163邮箱", "126邮箱", "Outlook", "Gmail", "自定义"])
        self.preset_combo.currentTextChanged.connect(self.update_preset_settings)
        layout.addRow("预设邮箱:", self.preset_combo)

        # 邮箱地址
        self.email_edit = QLineEdit()
        layout.addRow("邮箱地址:", self.email_edit)

        # 密码
        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        layout.addRow("密码:", self.password_edit)

        # 昵称
        self.nickname_edit = QLineEdit()
        layout.addRow("昵称(选填):", self.nickname_edit)

        # 协议选择
        self.protocol_label = QLabel("协议:")
        self.protocol_combo = QComboBox()
        self.protocol_combo.addItems(["IMAP", "POP3"])
        self.protocol_combo.currentTextChanged.connect(self.update_port_default)
        layout.addRow(self.protocol_label, self.protocol_combo)

        # 服务器地址
        self.server_label = QLabel("服务器:")
        self.server_edit = QLineEdit()
        layout.addRow(self.server_label, self.server_edit)

        # 端口
        self.port_label = QLabel("端口:")
        self.port_edit = QLineEdit()
        layout.addRow(self.port_label, self.port_edit)

        # SSL选项
        self.ssl_check = QCheckBox("使用SSL")
        self.ssl_check.setChecked(True)
        layout.addRow(self.ssl_check)


        # 按钮
        self.button_box = QHBoxLayout()
        self.ok_button = QPushButton("确定")
        self.ok_button.clicked.connect(self.accept)
        self.cancel_button = QPushButton("取消")
        self.cancel_button.clicked.connect(self.reject)
        self.button_box.addWidget(self.ok_button)
        self.button_box.addWidget(self.cancel_button)
        layout.addRow(self.button_box)

        # 初始化状态
        self.update_preset_settings("晏阳邮箱")

    def update_preset_settings(self, preset):
        preset_configs = {
            "晏阳邮箱": {
                "imap": {"server": "yanyn.cn", "port": "143", "ssl": False},
                "pop3": {"server": "yanyn.cn", "port": "110", "ssl": False}

            },
            "163邮箱": {
                "imap": {"server": "imap.163.com", "port": "993", "ssl": True},
                "pop3": {"server": "pop.163.com", "port": "995", "ssl": True}
            },
            "126邮箱": {
                "imap": {"server": "imap.126.com", "port": "993", "ssl": True},
                "pop3": {"server": "pop.126.com", "port": "995", "ssl": True}
            },
            "Outlook": {
                "imap": {"server": "outlook.office365.com", "port": "993", "ssl": True},
                "pop3": {"server": "outlook.office365.com", "port": "995", "ssl": True}
            },
            "Gmail": {
                "imap": {"server": "imap.gmail.com", "port": "993", "ssl": True},
                "pop3": {"server": "pop.gmail.com", "port": "995", "ssl": True}
            }
        }

        is_custom = preset == "自定义"

        self.protocol_label.setVisible(is_custom)
        self.protocol_combo.setVisible(is_custom)
        self.server_label.setVisible(is_custom)
        self.server_edit.setVisible(is_custom)
        self.port_label.setVisible(is_custom)
        self.port_edit.setVisible(is_custom)
        self.ssl_check.setVisible(is_custom)

        if not is_custom:
            protocol = self.protocol_combo.currentText().lower()
            config = preset_configs[preset][protocol]
            self.server_edit.setText(config["server"])
            self.port_edit.setText(str(config["port"]))
            self.ssl_check.setChecked(config["ssl"])
            self.protocol_combo.setCurrentText("IMAP")

    def update_port_default(self, protocol):
        if self.preset_combo.currentText() == "自定义":
            if protocol == "IMAP":
                self.port_edit.setText("993" if self.ssl_check.isChecked() else "143")
            else:  # POP3
                self.port_edit.setText("995" if self.ssl_check.isChecked() else "110")

    def get_account_info(self):
        return {
            "email": self.email_edit.text(),
            "password": self.password_edit.text(),
            "nickname": self.nickname_edit.text(),
            "protocol": self.protocol_combo.currentText(),
            "server": self.server_edit.text(),
            "port": int(self.port_edit.text()),
            "ssl": self.ssl_check.isChecked(),
            "preset": self.preset_combo.currentText()
        }


class EmailClient(QMainWindow):

    def cancel_operations(self):
        """安全地取消正在进行的操作"""
        with self.thread_lock:
            if self.worker_thread and self.worker_thread.is_alive():
                self.thread_cancel = True
                # 如果是IMAP操作，尝试优雅地关闭连接
                if hasattr(self, 'imap_conn') and self.imap_conn:
                    try:
                        self.imap_conn.logout()
                    except:
                        pass
                # 如果是POP3操作
                if hasattr(self, 'pop3_conn') and self.pop3_conn:
                    try:
                        self.pop3_conn.quit()
                    except:
                        pass

                # 等待线程结束，但不要无限等待
                self.worker_thread.join(timeout=2.0)
                if self.worker_thread.is_alive():
                    logging.warning("线程未能正常终止")

                self.thread_cancel = False
                self.worker_thread = None

    def fetch_emails_thread(self):
        try:
            with self.thread_lock:
                if self.thread_cancel:
                    return

                self.thread_running = True
                current_folder = self.folder_list.currentItem().text() if self.folder_list.currentItem() else "收件箱"
                folder_mapping = {
                    "收件箱": "INBOX",
                    "已发送": "Sent",
                    "草稿箱": "Drafts",
                    "垃圾邮件": "Trash"
                }
                folder = folder_mapping.get(current_folder, current_folder)

                if self.current_account["protocol"] == "IMAP":
                    self.imap_conn = imaplib.IMAP4_SSL(self.current_account["server"], self.current_account["port"]) if \
                        self.current_account["ssl"] else imaplib.IMAP4(self.current_account["server"],
                                                                       self.current_account["port"])
                    self.imap_conn.login(self.current_account["email"], self.current_account["password"])
                    self.imap_conn.select(folder)

                    status, messages = self.imap_conn.search(None, "ALL")
                    if status == "OK" and not self.thread_cancel:
                        emails = []
                        for email_id in reversed(messages[0].split()):
                            if self.thread_cancel:
                                break

                            status, msg_data = self.imap_conn.fetch(email_id, "(RFC822)")
                            if status == "OK":
                                emails.append(msg_data[0][1])

                        with QMutexLocker(self.email_mutex):
                            self.emails = emails
                        self.update_email_list_signal.emit_signal()

                    if not self.thread_cancel:
                        self.imap_conn.logout()
                    self.imap_conn = None

                else:  # POP3
                    self.pop3_conn = poplib.POP3_SSL(self.current_account["server"], self.current_account["port"]) if \
                        self.current_account["ssl"] else poplib.POP3(self.current_account["server"],
                                                                     self.current_account["port"])
                    self.pop3_conn.user(self.current_account["email"])
                    self.pop3_conn.pass_(self.current_account["password"])

                    if not self.thread_cancel:
                        emails = []
                        num_messages = len(self.pop3_conn.list()[1])
                        for i in range(num_messages, max(0, num_messages - 50), -1):
                            if self.thread_cancel:
                                break

                            response, lines, octets = self.pop3_conn.retr(i)
                            emails.append(b"\n".join(lines))

                        with QMutexLocker(self.email_mutex):
                            self.emails = emails
                        self.update_email_list_signal.emit_signal()

                    if not self.thread_cancel:
                        self.pop3_conn.quit()
                    self.pop3_conn = None

        except Exception as e:
            if not self.thread_cancel:  # 只有非取消导致的错误才记录
                logging.error(f"获取邮件失败: {str(e)}")
                self.status_bar.showMessage(f"错误: {str(e)}", 5000)
        finally:
            with self.thread_lock:
                self.thread_running = False
                self.worker_thread = None
                # 确保连接被关闭
                if hasattr(self, 'imap_conn') and self.imap_conn:
                    try:
                        self.imap_conn.logout()
                    except:
                        pass
                    self.imap_conn = None
                if hasattr(self, 'pop3_conn') and self.pop3_conn:
                    try:
                        self.pop3_conn.quit()
                    except:
                        pass
                    self.pop3_conn = None

    def apply_theme(self, theme_name):
        """应用指定的主题"""
        if theme_name == "azure":
            self.set_azure_theme()
        elif theme_name == "light":
            self.set_light_theme()
        else:  # dark
            self.set_dark_theme()

    def set_light_theme(self):
        """设置浅色主题"""
        palette = self.palette()
        palette.setColor(QPalette.ColorRole.Window, QColor(240, 240, 240))
        palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.black)
        palette.setColor(QPalette.ColorRole.Base, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.AlternateBase, QColor(240, 240, 240))
        palette.setColor(QPalette.ColorRole.ToolTipBase, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.ToolTipText, Qt.GlobalColor.black)
        palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.black)
        palette.setColor(QPalette.ColorRole.Button, QColor(240, 240, 240))
        palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.black)
        palette.setColor(QPalette.ColorRole.BrightText, Qt.GlobalColor.red)
        palette.setColor(QPalette.ColorRole.Highlight, QColor(0, 120, 215))
        palette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.white)
        self.setPalette(palette)

    def set_dark_theme(self):
        """设置深色主题"""
        palette = self.palette()
        palette.setColor(QPalette.ColorRole.Window, QColor(53, 53, 53))
        palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.Base, QColor(42, 42, 42))
        palette.setColor(QPalette.ColorRole.AlternateBase, QColor(53, 53, 53))
        palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(53, 53, 53))
        palette.setColor(QPalette.ColorRole.ToolTipText, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.Button, QColor(53, 53, 53))
        palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.BrightText, Qt.GlobalColor.red)
        palette.setColor(QPalette.ColorRole.Highlight, QColor(0, 120, 215))
        palette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.white)
        self.setPalette(palette)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Yanyn Email 0.14.1")
        self.setWindowIcon(QIcon("icon.ico"))
        self.resize(1200, 800)

        # 初始化线程相关属性
        self.worker_thread = None
        self.email_mutex = QMutex()
        self.update_email_list_signal = UpdateEmailListSignal()
        self.update_email_list_signal.signal.connect(self.update_email_list)

        # 初始化设置
        self.settings = QSettings("Yanyn", "Yanyn Email")
        self.accounts = self.settings.value("accounts", [])
        self.current_account = None
        self.emails = []

        # 设置蔚蓝色主题
        self.set_azure_theme()

        # 初始化UI
        self.init_ui()
        self.thread_lock = threading.Lock()
        self.thread_running = False
        self.thread_cancel = False
        self.thread_lock = threading.Lock()

        #系统托盘
        self.setup_tray_icon()

        # 加载保存的主题
        saved_theme = self.settings.value("theme", "azure")
        self.apply_theme(saved_theme)

        # 设置字体
        self.setFont(QFont("Microsoft YaHei", 10))

        # 现在可以安全地加载账户了
        if hasattr(self, 'account_list'):  # 确保 account_list 已初始化
            self.load_accounts()

        # 启动定时检查新邮件
        self.timer = QTimer()
        self.timer.timeout.connect(self.check_new_emails)
        self.timer.start(60000)  # 每分钟检查一次

    def set_azure_theme(self):
        palette = self.palette()
        azure_color = QColor(0, 127, 255)
        light_azure = QColor(200, 230, 255)
        dark_azure = QColor(0, 80, 160)

        palette.setColor(QPalette.ColorRole.Window, QColor(240, 248, 255))
        palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.black)
        palette.setColor(QPalette.ColorRole.Base, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.AlternateBase, light_azure)
        palette.setColor(QPalette.ColorRole.ToolTipBase, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.ToolTipText, Qt.GlobalColor.black)
        palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.black)
        palette.setColor(QPalette.ColorRole.Button, azure_color)
        palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.BrightText, Qt.GlobalColor.red)
        palette.setColor(QPalette.ColorRole.Highlight, dark_azure)
        palette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.white)

        self.setPalette(palette)
        self.setFont(QFont("Microsoft YaHei", 10))

    def init_ui(self):
        # 确保在创建新组件前清除旧组件
        if hasattr(self, 'account_list'):
            self.account_list.deleteLater()
        if hasattr(self, 'folder_list'):
            self.folder_list.deleteLater()
        if hasattr(self, 'email_list'):
            self.email_list.deleteLater()
        if hasattr(self, 'email_preview'):
            self.email_preview.deleteLater()

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        # 左侧面板
        left_panel = QWidget()
        left_panel.setMaximumWidth(250)
        left_layout = QVBoxLayout(left_panel)

        # 账户操作按钮
        account_buttons_layout = QHBoxLayout()
        self.add_account_btn = QPushButton("添加账户")
        self.add_account_btn.clicked.connect(self.add_account)
        self.delete_account_btn = QPushButton("删除账户")
        self.delete_account_btn.clicked.connect(self.delete_account)
        account_buttons_layout.addWidget(self.add_account_btn)
        account_buttons_layout.addWidget(self.delete_account_btn)

        # 账户列表
        self.account_list = QListWidget()
        self.account_list.itemClicked.connect(self.switch_account)

        # 文件夹列表
        self.folder_list = QListWidget()
        self.folder_list.itemClicked.connect(self.switch_folder)

        left_layout.addWidget(QLabel("邮箱账户"))
        left_layout.addWidget(self.account_list)
        left_layout.addLayout(account_buttons_layout)
        left_layout.addWidget(QLabel("邮箱文件夹"))
        left_layout.addWidget(self.folder_list)

        # 右侧面板
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)

        # 工具栏
        toolbar = QToolBar()
        self.refresh_btn = QAction(QIcon.fromTheme("view-refresh"), "刷新", self)
        self.refresh_btn.triggered.connect(self.refresh_emails)
        toolbar.addAction(self.refresh_btn)

        self.compose_btn = QAction(QIcon.fromTheme("mail-message-new"), "写邮件", self)
        self.compose_btn.triggered.connect(self.compose_email)
        toolbar.addAction(self.compose_btn)

        self.reply_btn = QAction(QIcon.fromTheme("mail-reply-sender"), "回复", self)
        self.reply_btn.triggered.connect(self.reply_email)
        toolbar.addAction(self.reply_btn)

        self.forward_btn = QAction(QIcon.fromTheme("mail-forward"), "转发", self)
        self.forward_btn.triggered.connect(self.forward_email)
        toolbar.addAction(self.forward_btn)

        self.delete_btn = QAction(QIcon.fromTheme("edit-delete"), "删除", self)
        self.delete_btn.triggered.connect(self.delete_email)
        toolbar.addAction(self.delete_btn)

        # 在工具栏添加设置按钮
        gear_icon = qta.icon("fa5s.cog")
        self.settings_btn = QAction(gear_icon, "设置", self)
        self.settings_btn.triggered.connect(self.show_settings)
        toolbar.addAction(self.settings_btn)

        right_layout.addWidget(toolbar)

        # 邮件列表和预览分割
        splitter = QSplitter(Qt.Orientation.Vertical)
        self.email_list = QTreeWidget()
        self.email_list.setHeaderLabels(["发件人", "主题", "日期"])
        self.email_list.setColumnWidth(0, 200)
        self.email_list.setColumnWidth(1, 300)
        self.email_list.setColumnWidth(2, 150)
        self.email_list.itemClicked.connect(self.show_email)

        # 邮件预览
        self.email_preview = QTabWidget()
        self.text_preview = QTextEdit()
        self.text_preview.setReadOnly(True)
        self.email_preview.addTab(self.text_preview, "文本")
        self.html_preview = QWebEngineView()
        self.email_preview.addTab(self.html_preview, "HTML")
        self.attachment_list = QListWidget()
        self.email_preview.addTab(self.attachment_list, "附件")

        splitter.addWidget(self.email_list)
        splitter.addWidget(self.email_preview)
        splitter.setSizes([300, 400])
        right_layout.addWidget(splitter)

        main_layout.addWidget(left_panel)
        main_layout.addWidget(right_panel)

        # 状态栏
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

    def show_settings(self):
        dialog = SettingsDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            settings = dialog.get_settings()
            self.settings.setValue("theme", settings["theme"])
            self.settings.setValue("minimize_to_tray", settings["minimize_to_tray"])

            # 应用新主题
            self.apply_theme(settings["theme"])

            QMessageBox.information(self, "成功", "设置已保存，部分设置可能需要重启应用才能生效")

    def load_accounts(self):
        self.account_list.clear()
        for account in self.accounts:
            self.account_list.addItem(f"{account.get('nickname', '未命名')} <{account['email']}>")

        if self.accounts:
            self.account_list.setCurrentRow(0)
            self.switch_account()

    def add_account(self):
        dialog = AddAccountDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            account_info = dialog.get_account_info()
            if not self.validate_account(account_info):
                return
            self.accounts.append(account_info)
            self.settings.setValue("accounts", self.accounts)
            self.load_accounts()

    def delete_account(self):
        selected_row = self.account_list.currentRow()
        if selected_row < 0 or selected_row >= len(self.accounts):
            QMessageBox.warning(self, "警告", "请先选择要删除的账户")
            return

        account_email = self.accounts[selected_row]['email']
        reply = QMessageBox.question(
            self, "确认删除",
            f"确定要删除账户 {account_email} 吗?\n此操作不可撤销!",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.No:
            return

        try:
            del self.accounts[selected_row]
            self.settings.setValue("accounts", self.accounts)

            if self.current_account and self.current_account['email'] == account_email:
                self.current_account = None
                self.folder_list.clear()
                self.email_list.clear()
                self.text_preview.clear()
                self.html_preview.setHtml("")
                self.attachment_list.clear()

            self.load_accounts()
            QMessageBox.information(self, "成功", "账户已删除")
        except Exception as e:
            logging.error(f"删除账户失败: {str(e)}")
            QMessageBox.critical(self, "错误", f"删除账户失败:\n{str(e)}")
            self.load_accounts()

    def validate_account(self, account_info):
        try:
            if account_info["protocol"] == "IMAP":
                mail = imaplib.IMAP4_SSL(account_info["server"], account_info["port"]) if account_info[
                    "ssl"] else imaplib.IMAP4(account_info["server"], account_info["port"])
                mail.login(account_info["email"], account_info["password"])
                mail.logout()
            else:  # POP3
                mail = poplib.POP3_SSL(account_info["server"], account_info["port"]) if account_info[
                    "ssl"] else poplib.POP3(account_info["server"], account_info["port"])
                mail.user(account_info["email"])
                mail.pass_(account_info["password"])
                mail.quit()
            return True
        except Exception as e:
            logging.error(f"账户验证失败: {str(e)}")
            QMessageBox.critical(self, "错误", f"无法连接到邮箱服务器:\n{str(e)}")
            return False

    def switch_account(self):
        # 先取消任何正在运行的线程
        self.cancel_operations()

        selected_row = self.account_list.currentRow()
        if 0 <= selected_row < len(self.accounts):
            self.current_account = self.accounts[selected_row]
            try:
                self.update_folder_list()
                self.refresh_emails()
            except Exception as e:
                logging.error(f"切换账户错误: {str(e)}")
                QMessageBox.critical(self, "错误", f"切换账户时发生错误:\n{str(e)}")

    def update_folder_list(self):
        self.folder_list.clear()
        if not self.current_account:
            self.status_bar.showMessage("未选择邮箱账户", 3000)
            return

        try:
            if self.current_account["protocol"] == "IMAP":
                if self.current_account["ssl"]:
                    mail = imaplib.IMAP4_SSL(self.current_account["server"], self.current_account["port"])
                else:
                    mail = imaplib.IMAP4(self.current_account["server"], self.current_account["port"])

                mail.login(self.current_account["email"], self.current_account["password"])
                status, folders = mail.list()
                mail.logout()

                if status == "OK":
                    # 文件夹名称映射为中文
                    folder_mapping = {
                        "INBOX": "收件箱",
                        "Sent": "已发送",
                        "Drafts": "草稿箱",
                        "Trash": "垃圾邮件",
                        "Junk": "垃圾邮件",
                        "Spam": "垃圾邮件",
                        "Archive": "存档"
                    }

                    for folder in folders:
                        folder_info = folder.decode()
                        if '"/"' in folder_info:
                            folder_name = folder_info.split('"/"')[-1].strip('"')
                        else:
                            folder_name = folder_info.split()[-1].strip('"')

                        # 使用中文名称，如果没有映射则使用原名
                        display_name = folder_mapping.get(folder_name, folder_name)
                        if folder_name:
                            self.folder_list.addItem(display_name)
            else:  # POP3协议
                self.folder_list.addItem("收件箱")

        except Exception as e:
            logging.error(f"更新文件夹列表失败: {str(e)}")
            self.status_bar.showMessage(f"错误: {str(e)}", 5000)

    def switch_folder(self):
        # 先取消任何正在运行的线程
        self.cancel_operations()
        try:
            self.refresh_emails()
        except Exception as e:
            logging.error(f"切换文件夹错误: {str(e)}")
            QMessageBox.critical(self, "错误", f"切换文件夹时发生错误:\n{str(e)}")

    def refresh_emails(self):
        with self.thread_lock:
            if not self.current_account or (self.worker_thread and self.worker_thread.isRunning()):
                return

            self.worker_thread = threading.Thread(target=self.fetch_emails_thread)
            self.worker_thread.start()

    def fetch_emails_thread(self):
        try:
            # 获取当前文件夹名称，并映射回英文名称用于IMAP操作
            current_folder = self.folder_list.currentItem().text() if self.folder_list.currentItem() else "收件箱"
            folder_mapping = {
                "收件箱": "INBOX",
                "已发送": "Sent",
                "草稿箱": "Drafts",
                "垃圾邮件": "Trash"
            }
            folder = folder_mapping.get(current_folder, current_folder)

            if self.current_account["protocol"] == "IMAP":
                mail = imaplib.IMAP4_SSL(self.current_account["server"], self.current_account["port"]) if \
                self.current_account["ssl"] else imaplib.IMAP4(self.current_account["server"],
                                                               self.current_account["port"])
                mail.login(self.current_account["email"], self.current_account["password"])
                mail.select(folder)

                status, messages = mail.search(None, "ALL")
                if status == "OK":
                    emails = []
                    for email_id in reversed(messages[0].split()):
                        status, msg_data = mail.fetch(email_id, "(RFC822)")
                        if status == "OK":
                            emails.append(msg_data[0][1])

                    with QMutexLocker(self.email_mutex):
                        self.emails = emails
                    self.update_email_list_signal.emit_signal()

                mail.logout()
            else:  # POP3
                mail = poplib.POP3_SSL(self.current_account["server"], self.current_account["port"]) if \
                self.current_account["ssl"] else poplib.POP3(self.current_account["server"],
                                                             self.current_account["port"])
                mail.user(self.current_account["email"])
                mail.pass_(self.current_account["password"])

                emails = []
                num_messages = len(mail.list()[1])
                for i in range(num_messages, max(0, num_messages - 50), -1):
                    response, lines, octets = mail.retr(i)
                    emails.append(b"\n".join(lines))

                with QMutexLocker(self.email_mutex):
                    self.emails = emails
                self.update_email_list_signal.emit_signal()
                mail.quit()
        except Exception as e:
            logging.error(f"获取邮件失败: {str(e)}")
            self.status_bar.showMessage(f"错误: {str(e)}", 5000)
        finally:
            with self.thread_lock:
                self.worker_thread = None

    def update_email_list(self):
        self.email_list.clear()

        with QMutexLocker(self.email_mutex):
            for raw_email in self.emails:
                try:
                    msg = email.message_from_bytes(raw_email)
                    from_ = parseaddr(msg.get("From", ""))[1] or "未知发件人"

                    subject, encoding = decode_header(msg.get("Subject", ""))[0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(encoding if encoding else "utf-8", errors="replace")
                    subject = subject or "(无主题)"

                    date = msg.get("Date", "未知日期")

                    item = QTreeWidgetItem([from_, subject, date])
                    item.setData(0, Qt.ItemDataRole.UserRole, raw_email)
                    self.email_list.addTopLevelItem(item)
                except Exception as e:
                    logging.error(f"解析邮件失败: {str(e)}")
                    continue

        if self.email_list.topLevelItemCount() > 0:
            self.email_list.setCurrentItem(self.email_list.topLevelItem(0))
            self.show_email()

    def show_email(self):
        item = self.email_list.currentItem()
        if not item:
            return

        raw_email = item.data(0, Qt.ItemDataRole.UserRole)
        msg = email.message_from_bytes(raw_email)

        # 获取HTML内容
        html_content = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/html":
                    payload = part.get_payload(decode=True)
                    charset = part.get_content_charset() or "utf-8"
                    html_content += payload.decode(charset, errors="replace")
        elif msg.get_content_type() == "text/html":
            payload = msg.get_payload(decode=True)
            charset = msg.get_content_charset() or "utf-8"
            html_content = payload.decode(charset, errors="replace")

        # 如果没有HTML内容，则使用纯文本转换为HTML
        if not html_content:
            text_content = self.get_text_content(msg)
            html_content = f"<pre style='font-family: sans-serif; white-space: pre-wrap;'>{text_content}</pre>"

        self.html_preview.setHtml(html_content)
        self.email_preview.setCurrentIndex(1)  # 切换到HTML标签页
        self.text_preview.setPlainText(self.html_to_plaintext(html_content))

        # 显示附件
        self.attachment_list.clear()
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_filename():
                    self.attachment_list.addItem(part.get_filename())

    def get_text_content(self, msg):
        text_content = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    payload = part.get_payload(decode=True)
                    charset = part.get_content_charset() or "utf-8"
                    text_content += payload.decode(charset, errors="replace")
        else:
            payload = msg.get_payload(decode=True)
            charset = msg.get_content_charset() or "utf-8"
            text_content = payload.decode(charset, errors="replace")
        return text_content

    def html_to_plaintext(self, html):
        text = re.sub('<[^<]+?>', '', html)
        text = re.sub(r'\n\s*\n', '\n\n', text)
        return text.strip()

    def compose_email(self):
        if not self.current_account:
            QMessageBox.warning(self, "警告", "请先选择发件邮箱账户")
            return

        dialog = ComposeEmailDialog(self)
        dialog.set_from_email(self.current_account["email"])

        if dialog.exec() == QDialog.DialogCode.Accepted:
            email_data = dialog.get_email_data()
            self.send_email(email_data)

    def send_email(self, email_data):
        try:
            if not self.current_account:
                QMessageBox.warning(self, "警告", "请先选择发件邮箱账户")
                return

            msg = MIMEMultipart()
            msg["From"] = self.current_account["email"]
            msg["To"] = email_data["to"]
            msg["Subject"] = email_data["subject"]

            text_part = MIMEText(email_data["content"], "plain", "utf-8")
            html_part = MIMEText(email_data["content"], "html", "utf-8")
            msg.attach(text_part)
            msg.attach(html_part)

            if email_data["attachment"]:
                with open(email_data["attachment"], "rb") as f:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition",
                                    f"attachment; filename={os.path.basename(email_data['attachment'])}")
                    msg.attach(part)

            smtp_server = self.get_smtp_server(self.current_account)
            with smtplib.SMTP_SSL(smtp_server["host"], smtp_server["port"]) if smtp_server["ssl"] else smtplib.SMTP(
                    smtp_server["host"], smtp_server["port"]) as server:
                if not smtp_server["ssl"] and smtp_server["port"] == 587:
                    server.starttls()
                server.login(self.current_account["email"], self.current_account["password"])
                server.send_message(msg)

            QMessageBox.information(self, "成功", "邮件发送成功")
        except Exception as e:
            logging.error(f"发送邮件失败: {str(e)}")
            QMessageBox.critical(self, "错误", f"发送邮件失败:\n{str(e)}")

    def get_smtp_server(self, account):
        preset_configs = {
            "晏阳邮箱": {"host": "yanyn.cn", "port": 25, "ssl": False},
            "QQ邮箱": {"host": "smtp.qq.com", "port": 465, "ssl": True},
            "163邮箱": {"host": "smtp.163.com", "port": 465, "ssl": True},
            "126邮箱": {"host": "smtp.126.com", "port": 465, "ssl": True},
            "Outlook": {"host": "smtp.office365.com", "port": 587, "ssl": True},
            "Gmail": {"host": "smtp.gmail.com", "port": 587, "ssl": True}
        }

        return preset_configs.get(account.get("preset"), {
            "host": account["server"],
            "port": account["port"],
            "ssl": account["ssl"]
        })

    def reply_email(self):
        item = self.email_list.currentItem()
        if not item:
            QMessageBox.warning(self, "警告", "请先选择要回复的邮件")
            return

        raw_email = item.data(0, Qt.ItemDataRole.UserRole)
        msg = email.message_from_bytes(raw_email)

        dialog = ComposeEmailDialog(self)
        dialog.setWindowTitle("回复邮件")
        dialog.set_from_email(self.current_account["email"])
        dialog.to_edit.setText(parseaddr(msg.get("From", ""))[1])
        dialog.subject_edit.setText(f"Re: {msg.get('Subject', '')}")

        text_content = self.get_text_content(msg)
        dialog.content_edit.setPlainText(f"\n\n---------- 原邮件 ----------\n{text_content}")

        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.send_email(dialog.get_email_data())

    def forward_email(self):
        item = self.email_list.currentItem()
        if not item:
            QMessageBox.warning(self, "警告", "请先选择要转发的邮件")
            return

        raw_email = item.data(0, Qt.ItemDataRole.UserRole)
        msg = email.message_from_bytes(raw_email)

        dialog = ComposeEmailDialog(self)
        dialog.setWindowTitle("转发邮件")
        dialog.set_from_email(self.current_account["email"])
        dialog.subject_edit.setText(f"Fwd: {msg.get('Subject', '')}")

        text_content = self.get_text_content(msg)
        dialog.content_edit.setPlainText(f"\n\n---------- 转发邮件 ----------\n{text_content}")

        # 添加附件
        if msg.is_multipart():
            temp_dir = "temp_attachments"
            os.makedirs(temp_dir, exist_ok=True)

            for part in msg.walk():
                if part.get_filename():
                    filename = part.get_filename()
                    filepath = os.path.join(temp_dir, filename)
                    with open(filepath, "wb") as f:
                        f.write(part.get_payload(decode=True))

                    dialog.attachment_path = filepath
                    dialog.attachment_label.setText(filename)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.send_email(dialog.get_email_data())

            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)

    def delete_email(self):
        item = self.email_list.currentItem()
        if not item:
            QMessageBox.warning(self, "警告", "请先选择要删除的邮件")
            return

        reply = QMessageBox.question(
            self, "确认删除",
            "确定要删除这封邮件吗?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.No:
            return

        try:
            if not self.current_account:
                QMessageBox.warning(self, "警告", "没有选中的邮箱账户")
                return

            raw_email = item.data(0, Qt.ItemDataRole.UserRole)
            msg = email.message_from_bytes(raw_email)
            message_id = msg.get("Message-ID", "")

            if self.current_account["protocol"] == "IMAP":
                mail = imaplib.IMAP4_SSL(self.current_account["server"], self.current_account["port"]) if \
                self.current_account["ssl"] else imaplib.IMAP4(self.current_account["server"],
                                                               self.current_account["port"])
                mail.login(self.current_account["email"], self.current_account["password"])

                folder = self.folder_list.currentItem().text() if self.folder_list.currentItem() else "INBOX"
                mail.select(folder)

                status, messages = mail.search(None, f'(HEADER Message-ID "{message_id}")')
                if status == "OK" and messages[0]:
                    mail.store(messages[0].split()[0], '+FLAGS', '\\Deleted')
                    mail.expunge()

                mail.logout()
            else:  # POP3
                mail = poplib.POP3_SSL(self.current_account["server"], self.current_account["port"]) if \
                self.current_account["ssl"] else poplib.POP3(self.current_account["server"],
                                                             self.current_account["port"])
                mail.user(self.current_account["email"])
                mail.pass_(self.current_account["password"])
                mail.dele(len(mail.list()[1]))
                mail.quit()

            self.email_list.takeTopLevelItem(self.email_list.indexOfTopLevelItem(item))
            QMessageBox.information(self, "成功", "邮件已删除")
        except Exception as e:
            logging.error(f"删除邮件失败: {str(e)}")
            QMessageBox.critical(self, "错误", f"删除邮件失败:\n{str(e)}")

    def check_new_emails(self):
        if self.current_account:
            self.refresh_emails()

    def closeEvent(self, event):
        """重写关闭事件，支持最小化到托盘"""
        if self.settings.value("minimize_to_tray", True, type=bool):
            event.ignore()
            self.hide()
            self.tray_icon.showMessage(
                "Yanyn Email",
                "应用程序已最小化到系统托盘",
                QIcon("icon.ico"),
                2000
            )
        else:
            self.settings.sync()
            event.accept()

    def setup_tray_icon(self):
        """设置系统托盘图标"""
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon("icon.ico"))

        # 创建托盘菜单
        tray_menu = QMenu()

        show_action = QAction("显示窗口", self)
        show_action.triggered.connect(self.show)

        exit_action = QAction("退出", self)
        exit_action.triggered.connect(self.close)

        tray_menu.addAction(show_action)
        tray_menu.addSeparator()
        tray_menu.addAction(exit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    client = EmailClient()
    client.show()
    sys.exit(app.exec())



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
                             QFileDialog, QComboBox, QCheckBox)
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
        self.preset_combo.addItems(["晏阳邮箱", "QQ邮箱", "163邮箱", "126邮箱", "Outlook", "Gmail", "自定义"])
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
            "QQ邮箱": {
                "imap": {"server": "imap.qq.com", "port": "993", "ssl": True},
                "pop3": {"server": "pop.qq.com", "port": "995", "ssl": True}
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
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Yanyn Email 0.14.0")
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

        # 加载账户
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
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        # 左侧面板
        left_panel = QWidget()
        left_panel.setMaximumWidth(250)
        left_layout = QVBoxLayout(left_panel)

        # 账户列表
        self.account_list = QListWidget()
        self.account_list.itemClicked.connect(self.switch_account)

        # 文件夹列表
        self.folder_list = QListWidget()
        self.folder_list.itemClicked.connect(self.switch_folder)

        # 账户操作按钮
        account_buttons_layout = QHBoxLayout()
        self.add_account_btn = QPushButton("添加账户")
        self.add_account_btn.clicked.connect(self.add_account)
        self.delete_account_btn = QPushButton("删除账户")
        self.delete_account_btn.clicked.connect(self.delete_account)
        account_buttons_layout.addWidget(self.add_account_btn)
        account_buttons_layout.addWidget(self.delete_account_btn)

        left_layout.addWidget(QLabel("账户"))
        left_layout.addWidget(self.account_list)
        left_layout.addWidget(QLabel("文件夹"))
        left_layout.addWidget(self.folder_list)
        left_layout.addLayout(account_buttons_layout)

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
        selected_row = self.account_list.currentRow()
        if 0 <= selected_row < len(self.accounts):
            self.current_account = self.accounts[selected_row]
            self.update_folder_list()
            self.refresh_emails()

    def update_folder_list(self):
        self.folder_list.clear()
        if not self.current_account:
            return

        try:
            if self.current_account["protocol"] == "IMAP":
                mail = imaplib.IMAP4_SSL(self.current_account["server"], self.current_account["port"]) if \
                self.current_account["ssl"] else imaplib.IMAP4(self.current_account["server"],
                                                               self.current_account["port"])
                mail.login(self.current_account["email"], self.current_account["password"])
                status, folders = mail.list()
                mail.logout()

                if status == "OK":
                    for folder in folders:
                        folder_name = folder.decode().split('"')[-2]
                        self.folder_list.addItem(folder_name)
            else:  # POP3
                self.folder_list.addItem("收件箱")
        except Exception as e:
            logging.error(f"更新文件夹列表失败: {str(e)}")
            self.status_bar.showMessage(f"错误: {str(e)}", 5000)

    def switch_folder(self):
        self.refresh_emails()

    def refresh_emails(self):
        if not self.current_account or (self.worker_thread and self.worker_thread.isRunning()):
            return

        self.worker_thread = threading.Thread(target=self.fetch_emails_thread)
        self.worker_thread.start()

    def fetch_emails_thread(self):
        try:
            folder = self.folder_list.currentItem().text() if self.folder_list.currentItem() else "INBOX"

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
        self.settings.sync()
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    client = EmailClient()
    client.show()
    sys.exit(app.exec())
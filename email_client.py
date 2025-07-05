import sys
import os
import logging
import socket
import imaplib
import poplib
import email
import smtplib
from email.message import EmailMessage
from email.header import decode_header
from email.utils import parseaddr
import time
import threading

import win32event
import winerror
import win32api
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


class EmailAccountDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("写邮件")
        self.setWindowIcon(QIcon("icon.ico"))
        self.resize(600, 500)

        layout = QVBoxLayout(self)

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
        # 预设邮箱的配置信息
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

        # 是否为自定义选项
        is_custom = preset == "自定义"

        # 设置字段的可见性
        self.protocol_label.setVisible(is_custom)
        self.protocol_combo.setVisible(is_custom)
        self.server_label.setVisible(is_custom)
        self.server_edit.setVisible(is_custom)
        self.port_label.setVisible(is_custom)
        self.port_edit.setVisible(is_custom)
        self.ssl_check.setVisible(is_custom)

        if not is_custom:
            # 获取当前协议
            protocol = self.protocol_combo.currentText().lower()

            # 更新配置
            config = preset_configs[preset][protocol]
            self.server_edit.setText(config["server"])
            self.port_edit.setText(str(config["port"]))
            self.ssl_check.setChecked(config["ssl"])

            # 锁定协议选择为IMAP（因为大多数预设只支持IMAP）
            self.protocol_combo.setCurrentText("IMAP")

    def update_port_default(self, protocol):
        # 只有在自定义模式下才更新端口
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
        # 必须首先调用父类初始化
        super().__init__()

        # 然后才能进行其他设置
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

        # 设置蔚蓝色调色板
        azure_color = QColor(0, 127, 255)
        light_azure = QColor(200, 230, 255)
        dark_azure = QColor(0, 80, 160)

        palette.setColor(QPalette.ColorRole.Window, QColor(240, 248, 255))  # AliceBlue背景
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

        # 设置字体
        font = QFont("Microsoft YaHei", 10)
        self.setFont(font)

    def init_ui(self):
        # 主窗口布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QHBoxLayout(central_widget)

        # 左侧账户和文件夹列表
        left_panel = QWidget()
        left_panel.setMaximumWidth(250)
        left_layout = QVBoxLayout(left_panel)

        # 账户列表
        self.account_list = QListWidget()
        self.account_list.itemClicked.connect(self.switch_account)
        left_layout.addWidget(QLabel("账户"))
        left_layout.addWidget(self.account_list)

        # 文件夹列表
        self.folder_list = QListWidget()
        self.folder_list.itemClicked.connect(self.switch_folder)
        left_layout.addWidget(QLabel("文件夹"))
        left_layout.addWidget(self.folder_list)

        # 添加账户按钮
        self.add_account_btn = QPushButton("添加账户")
        self.add_account_btn.clicked.connect(self.add_account)
        left_layout.addWidget(self.add_account_btn)

        main_layout.addWidget(left_panel)

        # 右侧主区域
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

        # 邮件列表
        self.email_list = QTreeWidget()
        self.email_list.setHeaderLabels(["发件人", "主题", "日期"])
        self.email_list.setColumnWidth(0, 200)
        self.email_list.setColumnWidth(1, 300)
        self.email_list.setColumnWidth(2, 150)
        self.email_list.itemClicked.connect(self.show_email)
        splitter.addWidget(self.email_list)

        # 邮件预览
        self.email_preview = QTabWidget()

        # 纯文本预览
        self.text_preview = QTextEdit()
        self.text_preview.setReadOnly(True)
        self.email_preview.addTab(self.text_preview, "文本")

        # HTML预览
        self.html_preview = QWebEngineView()
        self.email_preview.addTab(self.html_preview, "HTML")

        # 附件列表
        self.attachment_list = QListWidget()
        self.email_preview.addTab(self.attachment_list, "附件")

        splitter.addWidget(self.email_preview)
        splitter.setSizes([300, 400])

        right_layout.addWidget(splitter)

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
        dialog = EmailAccountDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            account_info = dialog.get_account_info()

            # 验证账户
            if not self.validate_account(account_info):
                return

            # 添加到账户列表
            self.accounts.append(account_info)
            self.settings.setValue("accounts", self.accounts)
            self.load_accounts()

    def validate_account(self, account_info):
        try:
            if account_info["protocol"] == "IMAP":
                if account_info["ssl"]:
                    mail = imaplib.IMAP4_SSL(account_info["server"], account_info["port"])
                else:
                    mail = imaplib.IMAP4(account_info["server"], account_info["port"])

                mail.login(account_info["email"], account_info["password"])
                mail.logout()
            else:  # POP3
                if account_info["ssl"]:
                    mail = poplib.POP3_SSL(account_info["server"], account_info["port"])
                else:
                    mail = poplib.POP3(account_info["server"], account_info["port"])

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
        if selected_row >= 0 and selected_row < len(self.accounts):
            self.current_account = self.accounts[selected_row]
            self.update_folder_list()
            self.refresh_emails()

    def update_folder_list(self):
        self.folder_list.clear()
        if not self.current_account:
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
        if not self.current_account:
            return

        # 使用工作线程获取邮件
        if self.worker_thread and self.worker_thread.isRunning():
            return

        self.worker_thread = threading.Thread(target=self.fetch_emails_thread)
        self.worker_thread.start()

    def fetch_emails_thread(self):
        try:
            folder = self.folder_list.currentItem().text() if self.folder_list.currentItem() else "INBOX"

            if self.current_account["protocol"] == "IMAP":
                if self.current_account["ssl"]:
                    mail = imaplib.IMAP4_SSL(self.current_account["server"], self.current_account["port"])
                else:
                    mail = imaplib.IMAP4(self.current_account["server"], self.current_account["port"])

                mail.login(self.current_account["email"], self.current_account["password"])
                mail.select(folder)

                status, messages = mail.search(None, "ALL")
                if status == "OK":
                    email_ids = messages[0].split()
                    emails = []

                    for email_id in reversed(email_ids):  # 从最新开始
                        status, msg_data = mail.fetch(email_id, "(RFC822)")
                        if status == "OK":
                            raw_email = msg_data[0][1]
                            emails.append(raw_email)

                    with QMutexLocker(self.email_mutex):
                        self.emails = emails

                    # 更新UI
                    self.update_email_list_signal.emit_signal()

                mail.logout()
            else:  # POP3
                if self.current_account["ssl"]:
                    mail = poplib.POP3_SSL(self.current_account["server"], self.current_account["port"])
                else:
                    mail = poplib.POP3(self.current_account["server"], self.current_account["port"])

                mail.user(self.current_account["email"])
                mail.pass_(self.current_account["password"])

                num_messages = len(mail.list()[1])
                emails = []

                for i in range(num_messages, max(0, num_messages - 50), -1):  # 获取最新的50封
                    response, lines, octets = mail.retr(i)
                    raw_email = b"\n".join(lines)
                    emails.append(raw_email)

                with QMutexLocker(self.email_mutex):
                    self.emails = emails

                # 更新UI
                self.update_email_list_signal.emit_signal()

                mail.quit()
        except Exception as e:
            logging.error(f"获取邮件失败: {str(e)}")
            self.status_bar.showMessage(f"错误: {str(e)}", 5000)
        finally:
            self.worker_thread = None  # 确保线程引用被清除

    def update_email_list(self):
        self.email_list.clear()

        with QMutexLocker(self.email_mutex):
            for raw_email in self.emails:
                try:
                    msg = email.message_from_bytes(raw_email)

                    # 解析发件人
                    from_ = parseaddr(msg.get("From", ""))[1]
                    if not from_:
                        from_ = "未知发件人"

                    # 解析主题
                    subject, encoding = decode_header(msg.get("Subject", ""))[0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(encoding if encoding else "utf-8", errors="replace")
                    if not subject:
                        subject = "(无主题)"

                    # 解析日期
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
                content_type = part.get_content_type()
                if content_type == "text/html":
                    payload = part.get_payload(decode=True)
                    charset = part.get_content_charset() or "utf-8"
                    html_content += payload.decode(charset, errors="replace")
        else:
            if msg.get_content_type() == "text/html":
                payload = msg.get_payload(decode=True)
                charset = msg.get_content_charset() or "utf-8"
                html_content = payload.decode(charset, errors="replace")

        # 如果没有HTML内容，则使用纯文本转换为HTML
        if not html_content:
            text_content = ""
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    if content_type == "text/plain":
                        payload = part.get_payload(decode=True)
                        charset = part.get_content_charset() or "utf-8"
                        text_content += payload.decode(charset, errors="replace")
            else:
                payload = msg.get_payload(decode=True)
                charset = msg.get_content_charset() or "utf-8"
                text_content = payload.decode(charset, errors="replace")

            # 将纯文本转换为HTML格式
            html_content = f"<pre style='font-family: sans-serif; white-space: pre-wrap;'>{text_content}</pre>"

        # 设置HTML为默认显示
        self.html_preview.setHtml(html_content)
        self.email_preview.setCurrentIndex(1)  # 切换到HTML标签页

        # 同时更新纯文本预览（保持同步）
        self.text_preview.setPlainText(self.html_to_plaintext(html_content))

        # 显示附件
        self.attachment_list.clear()
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_filename():
                    filename = part.get_filename()
                    self.attachment_list.addItem(filename)

    def html_to_plaintext(self, html):
        """简单将HTML转换为纯文本"""
        import re
        text = re.sub('<[^<]+?>', '', html)  # 移除HTML标签
        text = re.sub('\n\s*\n', '\n\n', text)  # 压缩多个空行
        return text.strip()

    def compose_email(self):
        dialog = EmailAccountDialog(self)  # 修改为正确的类名
        if dialog.exec() == QDialog.DialogCode.Accepted:
            email_data = dialog.get_email_data()
            self.send_email(email_data)

    def send_email(self, email_data):
        try:
            if not self.current_account:
                QMessageBox.warning(self, "警告", "请先选择发件邮箱账户")
                return

            # 创建MIME邮件
            msg = MIMEMultipart()
            msg["From"] = self.current_account["email"]
            msg["To"] = email_data["to"]
            msg["Subject"] = email_data["subject"]

            # 添加纯文本和HTML内容
            text_part = MIMEText(email_data["content"], "plain", "utf-8")
            html_part = MIMEText(email_data["content"], "html", "utf-8")
            msg.attach(text_part)
            msg.attach(html_part)

            # 添加附件
            if email_data["attachment"]:
                with open(email_data["attachment"], "rb") as f:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header(
                        "Content-Disposition",
                        f"attachment; filename={os.path.basename(email_data['attachment'])}",
                    )
                    msg.attach(part)

            # 发送邮件
            smtp_server = self.get_smtp_server(self.current_account)
            with smtplib.SMTP_SSL(smtp_server["host"], smtp_server["port"]) if smtp_server["ssl"] else smtplib.SMTP(
                    smtp_server["host"], smtp_server["port"]) as server:
                if not smtp_server["ssl"] and smtp_server["port"] == 587:  # STARTTLS
                    server.starttls()
                server.login(self.current_account["email"], self.current_account["password"])
                server.send_message(msg)

            QMessageBox.information(self, "成功", "邮件发送成功")
        except Exception as e:
            logging.error(f"发送邮件失败: {str(e)}")
            QMessageBox.critical(self, "错误", f"发送邮件失败:\n{str(e)}")

    def get_smtp_server(self, account):
        """获取SMTP服务器配置"""
        preset_configs = {
            "晏阳邮箱": {"host": "yanyn.cn", "port": 25, "ssl": False},
            "QQ邮箱": {"host": "smtp.qq.com", "port": 465, "ssl": True},
            "163邮箱": {"host": "smtp.163.com", "port": 465, "ssl": True},
            "126邮箱": {"host": "smtp.126.com", "port": 465, "ssl": True},
            "Outlook": {"host": "smtp.office365.com", "port": 587, "ssl": True},
            "Gmail": {"host": "smtp.gmail.com", "port": 587, "ssl": True}
        }

        if account.get("preset") in preset_configs:
            return preset_configs[account["preset"]]
        else:
            # 默认配置
            return {"host": account["server"], "port": account["port"], "ssl": account["ssl"]}

    def reply_email(self):
        item = self.email_list.currentItem()
        if not item:
            QMessageBox.warning(self, "警告", "请先选择要回复的邮件")
            return

        raw_email = item.data(0, Qt.ItemDataRole.UserRole)
        msg = email.message_from_bytes(raw_email)

        dialog = EmailAccountDialog(self)
        dialog.setWindowTitle("回复邮件")

        # 设置收件人
        from_ = parseaddr(msg.get("From", ""))[1]
        dialog.to_edit.setText(from_)

        # 设置主题
        subject = msg.get("Subject", "")
        dialog.subject_edit.setText(f"Re: {subject}")

        # 设置引用内容
        text_content = self.get_text_content(msg)
        quoted_content = f"\n\n---------- 原邮件 ----------\n{text_content}"
        dialog.content_edit.setPlainText(quoted_content)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            email_data = dialog.get_email_data()
            self.send_email(email_data)

    def forward_email(self):
        item = self.email_list.currentItem()
        if not item:
            QMessageBox.warning(self, "警告", "请先选择要转发的邮件")
            return

        raw_email = item.data(0, Qt.ItemDataRole.UserRole)
        msg = email.message_from_bytes(raw_email)

        dialog = EmailAccountDialog(self)
        dialog.setWindowTitle("转发邮件")

        # 设置主题
        subject = msg.get("Subject", "")
        dialog.subject_edit.setText(f"Fwd: {subject}")

        # 设置转发内容
        text_content = self.get_text_content(msg)
        forwarded_content = f"\n\n---------- 转发邮件 ----------\n{text_content}"
        dialog.content_edit.setPlainText(forwarded_content)

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
            email_data = dialog.get_email_data()
            self.send_email(email_data)

            # 清理临时附件
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
                if self.current_account["ssl"]:
                    mail = imaplib.IMAP4_SSL(self.current_account["server"], self.current_account["port"])
                else:
                    mail = imaplib.IMAP4(self.current_account["server"], self.current_account["port"])

                mail.login(self.current_account["email"], self.current_account["password"])

                folder = self.folder_list.currentItem().text() if self.folder_list.currentItem() else "INBOX"
                mail.select(folder)

                # 搜索并删除邮件
                status, messages = mail.search(None, f'(HEADER Message-ID "{message_id}")')
                if status == "OK" and messages[0]:
                    mail.store(messages[0].split()[0], '+FLAGS', '\\Deleted')
                    mail.expunge()

                mail.logout()
            else:  # POP3
                if self.current_account["ssl"]:
                    mail = poplib.POP3_SSL(self.current_account["server"], self.current_account["port"])
                else:
                    mail = poplib.POP3(self.current_account["server"], self.current_account["port"])

                mail.user(self.current_account["email"])
                mail.pass_(self.current_account["password"])

                # POP3通常不支持选择性删除，这里简单实现删除最新邮件
                num_messages = len(mail.list()[1])
                mail.dele(num_messages)

                mail.quit()

            # 从UI中移除
            self.email_list.takeTopLevelItem(self.email_list.indexOfTopLevelItem(item))
            QMessageBox.information(self, "成功", "邮件已删除")

        except Exception as e:
            logging.error(f"删除邮件失败: {str(e)}")
            QMessageBox.critical(self, "错误", f"删除邮件失败:\n{str(e)}")

    def get_text_content(self, msg):
        """获取邮件的纯文本内容"""
        text_content = ""
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == "text/plain":
                    payload = part.get_payload(decode=True)
                    charset = part.get_content_charset() or "utf-8"
                    text_content += payload.decode(charset, errors="replace")
        else:
            payload = msg.get_payload(decode=True)
            charset = msg.get_content_charset() or "utf-8"
            text_content = payload.decode(charset, errors="replace")

        return text_content

    def check_new_emails(self):
        # 定时检查新邮件
        if self.current_account:
            self.refresh_emails()

    def closeEvent(self, event):
        # 保存设置
        self.settings.sync()
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)

    # 设置应用程序样式
    app.setStyle("Fusion")

    client = EmailClient()
    client.show()

    sys.exit(app.exec())
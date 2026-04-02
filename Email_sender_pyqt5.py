import os
import sys
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.image import MIMEImage
from email.utils import formataddr
import openpyxl  # pyright: ignore[reportMissingModuleSource]
from PyQt5.QtGui import QIcon, QFont, QPixmap  # pyright: ignore[reportMissingImports]
from PyQt5.QtWidgets import (  # pyright: ignore[reportMissingImports]
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLineEdit, QPushButton,
    QTextEdit, QListWidget, QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem, QProgressBar,
    QGroupBox, QFormLayout, QCheckBox, QDesktopWidget
)
from PyQt5.QtCore import Qt  # pyright: ignore[reportMissingImports]

# 全局变量
email_list = []          # 收件人邮箱列表
attachment_paths = []    # 附件路径列表
sender_accounts = []     # 发件人账号池: [(email, password), ...]

def resource_path(relative_path):
    """获取资源的绝对路径，兼容 PyInstaller 打包后的环境"""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# 获取SMTP服务器地址
def get_smtp_server(email):
    if email.endswith('@qq.com'):
        return 'smtp.qq.com', 465
    elif email.endswith('@163.com'):
        return 'smtp.163.com', 465
    else:
        return None, None

# 发送邮件函数（增加 is_html 参数）
def send_email(sender_email, sender_nickname, sender_pass, recipient_email,
               subject, message, attachment_paths, is_html=False):
    msg = MIMEMultipart()
    msg['Subject'] = subject

    # 设置发件人信息
    display_info = sender_email
    if sender_nickname:
        msg['From'] = formataddr((sender_nickname, display_info))
    else:
        msg['From'] = display_info

    msg['To'] = recipient_email

    # 添加邮件正文（根据 is_html 决定类型）
    if is_html:
        msg.attach(MIMEText(message, 'html', 'utf-8'))
    else:
        msg.attach(MIMEText(message, 'plain', 'utf-8'))

    # 添加所有附件
    for attachment_path in attachment_paths:
        try:
            if attachment_path.lower().endswith(('.jpg', '.jpeg', '.png')):
                with open(attachment_path, 'rb') as attachment:
                    part = MIMEImage(attachment.read())
                    part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(attachment_path))
                    msg.attach(part)
            else:
                with open(attachment_path, 'rb') as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(attachment_path))
                    msg.attach(part)
        except Exception as e:
            show_custom_message_box(QMessageBox.Critical, f"添加附件 {attachment_path} 失败: {e}", "错误")

    smtp = None
    try:
        smtp_server, smtp_port = get_smtp_server(sender_email)
        if not smtp_server:
            show_custom_message_box(QMessageBox.Critical, "不支持的邮箱域名", "错误")
            return False, recipient_email
        smtp = smtplib.SMTP_SSL(smtp_server, smtp_port)
        smtp.login(sender_email, sender_pass)
        smtp.send_message(msg)
        return True, None
    except Exception as e:
        return False, recipient_email
    finally:
        if smtp:
            smtp.quit()

def show_custom_message_box(icon, text, title):
    msg_box = QMessageBox()
    msg_box.setIcon(icon)
    msg_box.setText(text)
    msg_box.setWindowTitle(title)

    if icon == QMessageBox.Information:
        pixmap = QPixmap(resource_path('success_icon.png'))
    elif icon == QMessageBox.Warning:
        pixmap = QPixmap(resource_path('warning_icon.png'))
    elif icon == QMessageBox.Critical:
        pixmap = QPixmap(resource_path('error_icon.png'))
    else:
        pixmap = QPixmap()
    msg_box.setIconPixmap(pixmap.scaled(64, 64, Qt.KeepAspectRatio))

    font = QFont("Arial", 14, QFont.Bold)
    msg_box.setFont(font)

    msg_box.setStyleSheet("""
        QMessageBox {
            background-color: #f8f9fa;
            border: 2px solid #ced4da;
            border-radius: 10px;
            padding: 20px;
        }
        QMessageBox QLabel {
            color: #212529;
            padding: 10px;
            font-size: 16px;
            text-align: center;
        }
        QMessageBox QPushButton {
            background-color: #007bff;
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 5px;
            font-size: 14px;
            min-width: 80px;
        }
        QMessageBox QPushButton:hover {
            background-color: #0056b3;
        }
        QMessageBox QPushButton:pressed {
            background-color: #004080;
        }
    """)

    msg_box.exec_()

class EmailSenderApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.is_logged_in = False
        self.initUI()

    def initUI(self):
        self.setWindowTitle("邮件发送工具 BY 寂寞沙洲冷 QQ/VX:1638276310")
        # 窗口居中：获取屏幕大小并计算位置
        screen = QDesktopWidget().screenGeometry()
        width = 1750
        height = 1300
        x = (screen.width() - width) // 2
        y = (screen.height() - height) // 2
        self.setGeometry(x, y, width, height)
        self.setWindowIcon(QIcon(resource_path('email_icon.png')))

        font = QFont("Arial", 10)
        QApplication.setFont(font)

        main_widget = QWidget()
        main_layout = QHBoxLayout()
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

        # ---------- 左侧布局 ----------
        left_widget = QWidget()
        left_layout = QVBoxLayout()
        left_widget.setLayout(left_layout)

        # 发件人信息组（单账号）
        sender_group = QGroupBox("发件人信息（单账号）")
        sender_layout = QFormLayout()
        self.entry_login_email = QLineEdit()
        self.entry_login_email.setPlaceholderText("请输入登录邮箱")
        self.entry_sender_nickname = QLineEdit()
        self.entry_sender_nickname.setPlaceholderText("请输入发件人昵称（可选）")
        self.entry_password = QLineEdit()
        self.entry_password.setPlaceholderText("请输入授权码")
        self.entry_password.setEchoMode(QLineEdit.Password)

        sender_layout.addRow("登录邮箱:", self.entry_login_email)
        sender_layout.addRow("发件人昵称:", self.entry_sender_nickname)
        sender_layout.addRow("授权码:", self.entry_password)
        sender_group.setLayout(sender_layout)
        left_layout.addWidget(sender_group)

        # 测试登录按钮
        self.button_test_login = QPushButton("测试SMTP登录")
        self.button_test_login.setStyleSheet("background-color: #4CAF50; color: white;")
        self.button_test_login.clicked.connect(self.test_smtp_login)
        left_layout.addWidget(self.button_test_login)

        # ---------- 多账号池组 ----------
        account_pool_group = QGroupBox("发件人账号池（多账号发送）")
        account_pool_layout = QVBoxLayout()

        # 上传/清空按钮
        btn_layout = QHBoxLayout()
        self.button_upload_accounts = QPushButton("上传账号列表")
        self.button_upload_accounts.setStyleSheet("background-color: #008CBA; color: white;")
        self.button_upload_accounts.clicked.connect(self.upload_accounts)
        self.button_clear_accounts = QPushButton("清空账号列表")
        self.button_clear_accounts.setStyleSheet("background-color: #f0ad4e; color: white;")
        self.button_clear_accounts.clicked.connect(self.clear_accounts)
        btn_layout.addWidget(self.button_upload_accounts)
        btn_layout.addWidget(self.button_clear_accounts)
        account_pool_layout.addLayout(btn_layout)

        # 账号列表显示
        self.listbox_accounts = QListWidget()
        account_pool_layout.addWidget(self.listbox_accounts)
        account_pool_group.setLayout(account_pool_layout)
        left_layout.addWidget(account_pool_group)

        # ---------- 发送模式选择 ----------
        self.checkbox_multi_mode = QCheckBox("启用多账号发送（平均分配）")
        self.checkbox_multi_mode.setChecked(False)  # 默认单账号
        left_layout.addWidget(self.checkbox_multi_mode)

        # 邮箱列表
        upload_clear_layout = QHBoxLayout()
        self.button_upload_email_list = QPushButton("上传邮箱列表")
        self.button_upload_email_list.setStyleSheet("background-color: #008CBA; color: white;")
        self.button_upload_email_list.clicked.connect(self.upload_email_list)
        self.button_clear_email_list = QPushButton("清空邮箱列表")
        self.button_clear_email_list.setStyleSheet("background-color: #f0ad4e; color: white;")
        self.button_clear_email_list.clicked.connect(self.clear_email_list)
        upload_clear_layout.addWidget(self.button_upload_email_list)
        upload_clear_layout.addWidget(self.button_clear_email_list)
        left_layout.addLayout(upload_clear_layout)

        email_list_group = QGroupBox("邮箱列表 (0)")
        email_list_layout = QVBoxLayout()
        self.listbox_email_list = QListWidget()
        email_list_layout.addWidget(self.listbox_email_list)
        email_list_group.setLayout(email_list_layout)
        left_layout.addWidget(email_list_group)
        self.email_list_group = email_list_group

        # 附件
        self.button_upload_attachment = QPushButton("上传附件")
        self.button_upload_attachment.setStyleSheet("background-color: #be14807f; color: black;")
        self.button_upload_attachment.clicked.connect(self.upload_attachment)
        left_layout.addWidget(self.button_upload_attachment)

        attachment_group = QGroupBox("附件列表")
        attachment_layout = QVBoxLayout()
        self.listbox_attachments = QListWidget()
        attachment_layout.addWidget(self.listbox_attachments)
        attachment_group.setLayout(attachment_layout)
        left_layout.addWidget(attachment_group)

        # 邮件内容组（增加HTML选项）
        message_group = QGroupBox("邮件内容")
        message_layout = QFormLayout()
        self.entry_subject = QLineEdit()
        self.entry_subject.setPlaceholderText("请输入邮件主题")
        self.text_message = QTextEdit()
        self.text_message.setPlaceholderText("请输入邮件正文（支持HTML代码）")
        self.checkbox_html = QCheckBox("正文为HTML格式")
        self.checkbox_html.setChecked(False)
        message_layout.addRow("邮件主题:", self.entry_subject)
        message_layout.addRow("邮件正文:", self.text_message)
        message_layout.addRow("", self.checkbox_html)
        message_group.setLayout(message_layout)
        left_layout.addWidget(message_group)

        # 发送按钮
        self.button_send_emails = QPushButton("发送邮件")
        self.button_send_emails.setStyleSheet("background-color: #008CBA; color: white;")
        self.button_send_emails.clicked.connect(self.send_emails)
        left_layout.addWidget(self.button_send_emails)

        # ---------- 右侧布局 ----------
        right_widget = QWidget()
        right_layout = QVBoxLayout()
        right_widget.setLayout(right_layout)

        status_group = QGroupBox("发送状态")
        status_layout = QVBoxLayout()
        self.table_status = QTableWidget()
        self.table_status.setColumnCount(2)
        self.table_status.setHorizontalHeaderLabels(["邮箱", "状态"])
        self.table_status.setColumnWidth(0, 300)
        self.table_status.setColumnWidth(1, 100)
        self.table_status.horizontalHeader().setStretchLastSection(True)
        status_layout.addWidget(self.table_status)
        status_group.setLayout(status_layout)
        right_layout.addWidget(status_group)

        self.progress_bar = QProgressBar()
        right_layout.addWidget(self.progress_bar)

        main_layout.addWidget(left_widget)
        main_layout.addWidget(right_widget)

        self.setStyleSheet("""
            QGroupBox {
                border: 1px solid gray;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px;
            }
            QPushButton {
                padding: 5px;
                border-radius: 5px;
            }
            QLineEdit, QTextEdit, QListWidget, QTableWidget {
                border: 1px solid #ccc;
                border-radius: 5px;
                padding: 5px;
            }
        """)

    # ---------- 邮箱列表管理 ----------
    def update_email_count(self):
        count = len(email_list)
        self.email_list_group.setTitle(f"邮箱列表 ({count})")

    def clear_email_list(self):
        global email_list
        email_list.clear()
        self.listbox_email_list.clear()
        self.update_email_count()
        show_custom_message_box(QMessageBox.Information, "邮箱列表已清空", "提示")

    def upload_email_list(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择邮箱列表文件",
            "",
            "All supported files (*.xlsx *.txt)"
        )
        if file_path:
            if file_path.endswith('.xlsx'):
                data = openpyxl.load_workbook(file_path)
                sheet = data.active
                column_values = []
                for cell in sheet.iter_rows(min_col=1, max_col=1, values_only=True):
                    if cell[0] is not None and not cell[0].endswith('@yahoo.com'):
                        column_values.append(cell[0])
                email_list.extend(column_values)
            elif file_path.endswith('.txt'):
                with open(file_path, 'r', encoding='utf-8') as f:
                    email_data = f.read().split('\n')
                    for email in email_data:
                        if email.strip() and not email.strip().startswith('#') and not email.strip().endswith('@yahoo.com'):
                            email_list.append(email.strip())
            self.listbox_email_list.clear()
            for email in email_list:
                self.listbox_email_list.addItem(email)
            self.update_email_count()
            show_custom_message_box(QMessageBox.Information, "邮箱列表上传成功！", "成功")

    # ---------- 附件管理 ----------
    def upload_attachment(self):
        file_paths, _ = QFileDialog.getOpenFileNames(self, "选择附件", "", "All files (*.*)")
        if file_paths:
            for file_path in file_paths:
                if os.path.getsize(file_path) > 10 * 1024 * 1024:
                    show_custom_message_box(QMessageBox.Warning, f"文件 {file_path} 超过10MB，跳过此附件。", "警告")
                    continue
                attachment_paths.append(file_path)
                self.listbox_attachments.addItem(file_path)
            show_custom_message_box(QMessageBox.Information, "附件上传成功！", "成功")

    # ---------- 账号池管理 ----------
    def upload_accounts(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择发件人账号列表",
            "",
            "Text files (*.txt)"
        )
        if not file_path:
            return
        global sender_accounts
        sender_accounts.clear()
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            for line in lines:
                line = line.strip()
                if not line or line.startswith('#'):
                    continue
                # 格式: 邮箱-授权码
                if '-' in line:
                    email, pwd = line.split('-', 1)
                    email = email.strip()
                    pwd = pwd.strip()
                    if email and pwd:
                        sender_accounts.append((email, pwd))
                else:
                    show_custom_message_box(QMessageBox.Warning, f"格式错误，跳过：{line}", "警告")
            self.listbox_accounts.clear()
            for email, _ in sender_accounts:
                self.listbox_accounts.addItem(email)
            show_custom_message_box(QMessageBox.Information, f"成功加载 {len(sender_accounts)} 个发件人账号", "成功")
        except Exception as e:
            show_custom_message_box(QMessageBox.Critical, f"读取文件失败: {e}", "错误")

    def clear_accounts(self):
        global sender_accounts
        sender_accounts.clear()
        self.listbox_accounts.clear()
        show_custom_message_box(QMessageBox.Information, "账号列表已清空", "提示")

    # ---------- 登录测试 ----------
    def test_smtp_login(self):
        login_email = self.entry_login_email.text()
        sender_pass = self.entry_password.text()
        smtp = None
        try:
            smtp_server, smtp_port = get_smtp_server(login_email)
            if not smtp_server:
                show_custom_message_box(QMessageBox.Critical, "不支持的邮箱域名", "错误")
                return
            smtp = smtplib.SMTP_SSL(smtp_server, smtp_port)
            smtp.login(login_email, sender_pass)
            self.is_logged_in = True
            show_custom_message_box(QMessageBox.Information, "SMTP登录成功！", "成功")
        except Exception as e:
            self.is_logged_in = False
            show_custom_message_box(QMessageBox.Critical, f"SMTP登录失败: {e}", "错误")
        finally:
            if smtp:
                smtp.quit()

    # ---------- 发送邮件 ----------
    def send_emails(self):
        global_nickname = self.entry_sender_nickname.text()
        subject = self.entry_subject.text()
        message = self.text_message.toPlainText()
        is_html = self.checkbox_html.isChecked()

        if not subject or not message:
            show_custom_message_box(QMessageBox.Warning, "请填写邮件主题和正文！", "警告")
            return

        # 判断发送模式
        multi_mode = self.checkbox_multi_mode.isChecked()

        if multi_mode:
            # 多账号模式：必须使用账号池
            if not sender_accounts:
                show_custom_message_box(QMessageBox.Warning, "多账号模式需要先上传账号列表！", "警告")
                return
            if not email_list:
                show_custom_message_box(QMessageBox.Warning, "请先上传收件人邮箱列表！", "警告")
                return

            # 平均分配收件人
            account_count = len(sender_accounts)
            recipient_count = len(email_list)
            base = recipient_count // account_count
            remainder = recipient_count % account_count
            slices = []
            start = 0
            for i in range(account_count):
                end = start + base + (1 if i < remainder else 0)
                slices.append(email_list[start:end])
                start = end

            success_count = 0
            failure_count = 0
            failure_emails = []

            total = recipient_count
            self.table_status.setRowCount(total)
            self.progress_bar.setMaximum(total)
            self.progress_bar.setValue(0)

            current_row = 0  # 用于填充状态表
            for idx, (account, recipient_slice) in enumerate(zip(sender_accounts, slices)):
                sender_email, sender_pass = account
                # 可选：显示当前使用的账号信息
                for recipient in recipient_slice:
                    success, failed_email = send_email(
                        sender_email, global_nickname, sender_pass,
                        recipient, subject, message, attachment_paths, is_html
                    )
                    if success:
                        success_count += 1
                        self.table_status.setItem(current_row, 0, QTableWidgetItem(recipient))
                        self.table_status.setItem(current_row, 1, QTableWidgetItem("成功"))
                    else:
                        failure_count += 1
                        failure_emails.append(failed_email)
                        self.table_status.setItem(current_row, 0, QTableWidgetItem(recipient))
                        self.table_status.setItem(current_row, 1, QTableWidgetItem("失败"))
                    self.progress_bar.setValue(current_row + 1)
                    QApplication.processEvents()
                    current_row += 1

            show_custom_message_box(
                QMessageBox.Information,
                f"成功发送 {success_count} 封邮件，失败 {failure_count} 封。",
                "发送结果"
            )
        else:
            # 单账号模式：使用界面上的账号
            login_email = self.entry_login_email.text()
            sender_pass = self.entry_password.text()
            if not login_email or not sender_pass:
                show_custom_message_box(QMessageBox.Warning, "请填写发件人信息！", "警告")
                return
            if not self.is_logged_in:
                show_custom_message_box(QMessageBox.Warning, "请先测试SMTP登录并成功登录！", "警告")
                return
            if not email_list:
                show_custom_message_box(QMessageBox.Warning, "请先上传收件人邮箱列表！", "警告")
                return

            success_count = 0
            failure_count = 0
            failure_emails = []

            total = len(email_list)
            self.table_status.setRowCount(total)
            self.progress_bar.setMaximum(total)
            self.progress_bar.setValue(0)

            for i, recipient in enumerate(email_list):
                success, failed_email = send_email(
                    login_email, global_nickname, sender_pass,
                    recipient, subject, message, attachment_paths, is_html
                )
                if success:
                    success_count += 1
                    self.table_status.setItem(i, 0, QTableWidgetItem(recipient))
                    self.table_status.setItem(i, 1, QTableWidgetItem("成功"))
                else:
                    failure_count += 1
                    failure_emails.append(failed_email)
                    self.table_status.setItem(i, 0, QTableWidgetItem(recipient))
                    self.table_status.setItem(i, 1, QTableWidgetItem("失败"))
                self.progress_bar.setValue(i + 1)
                QApplication.processEvents()

            show_custom_message_box(
                QMessageBox.Information,
                f"成功发送 {success_count} 封邮件，失败 {failure_count} 封。",
                "发送结果"
            )

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = EmailSenderApp()
    ex.show()
    sys.exit(app.exec_())
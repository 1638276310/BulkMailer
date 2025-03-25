import sys
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.image import MIMEImage
import openpyxl
import os
from PyQt5.QtGui import QIcon, QFont, QPixmap
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLineEdit, QPushButton,
    QTextEdit, QListWidget, QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem, QProgressBar,
    QGroupBox, QFormLayout
)
from PyQt5.QtCore import Qt

# 全局变量
email_list = []  # 邮箱列表
attachment_paths = []  # 附件路径列表

# 获取SMTP服务器地址
def get_smtp_server(email):
    if email.endswith('@qq.com'):
        return 'smtp.qq.com', 465
    elif email.endswith('@163.com'):
        return 'smtp.163.com', 465
    else:
        return None, None

# 发送邮件函数
def send_email(sender_email, sender_pass, recipient_email, subject, message, attachment_paths):
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = recipient_email

    # 添加邮件正文（纯文本）
    msg.attach(MIMEText(message, 'plain', 'utf-8'))

    # 添加所有附件
    for attachment_path in attachment_paths:
        try:
            if attachment_path.endswith('.jpg') or attachment_path.endswith('.jpeg') or attachment_path.endswith('.png'):
                with open(attachment_path, 'rb') as attachment:
                    part = MIMEImage(attachment.read())
                    part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(attachment_path))
                    msg.attach(part)
            elif attachment_path.endswith('.xlsx') or attachment_path.endswith('.xls'):
                with open(attachment_path, 'rb') as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(attachment_path))
                    msg.attach(part)
            elif attachment_path.endswith('.txt'):
                with open(attachment_path, 'rb') as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(attachment_path))
                    msg.attach(part)
            else:
                # 处理其他类型的文件
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

# 自定义消息框函数
def show_custom_message_box(icon, text, title):
    msg_box = QMessageBox()
    msg_box.setIcon(icon)
    msg_box.setText(text)
    msg_box.setWindowTitle(title)

    # 设置自定义图标
    if icon == QMessageBox.Information:
        pixmap = QPixmap('success_icon.png')  # 使用✅图标
    elif icon == QMessageBox.Warning:
        pixmap = QPixmap('warning_icon.png')  # 使用⚠️图标
    elif icon == QMessageBox.Critical:
        pixmap = QPixmap('error_icon.png')  # 使用❌图标
    else:
        pixmap = QPixmap()
    msg_box.setIconPixmap(pixmap.scaled(64, 64, Qt.KeepAspectRatio))  # 放大图标

    # 设置字体
    font = QFont("Arial", 14, QFont.Bold)
    msg_box.setFont(font)

    # 设置样式表
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
            text-align: center;  /* 文字居中 */
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

    # 强制刷新样式表
    msg_box.setStyleSheet(msg_box.styleSheet())

    msg_box.exec_()

class EmailSenderApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.is_logged_in = False  # 登录状态，初始为未登录
        self.initUI()

    def initUI(self):
        self.setWindowTitle("邮件发送工具 BY 寂寞沙洲冷 QV1638276310")
        self.setGeometry(300, 200, 1200, 600)
        self.setWindowIcon(QIcon('email_icon.png'))  # 设置窗口图标

        # 设置全局字体
        font = QFont("Arial", 10)
        QApplication.setFont(font)

        # 主布局
        main_widget = QWidget()
        main_layout = QHBoxLayout()
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

        # 左侧布局
        left_widget = QWidget()
        left_layout = QVBoxLayout()
        left_widget.setLayout(left_layout)

        # 发件人邮箱和授权码
        sender_group = QGroupBox("发件人信息")
        sender_layout = QFormLayout()
        self.entry_sender_email = QLineEdit()
        self.entry_sender_email.setPlaceholderText("请输入发件人邮箱")
        self.entry_password = QLineEdit()
        self.entry_password.setPlaceholderText("请输入授权码")
        self.entry_password.setEchoMode(QLineEdit.Password)
        sender_layout.addRow("发件人邮箱:", self.entry_sender_email)
        sender_layout.addRow("授权码:", self.entry_password)
        sender_group.setLayout(sender_layout)
        left_layout.addWidget(sender_group)

        # 测试SMTP登录按钮
        self.button_test_login = QPushButton("测试SMTP登录")
        self.button_test_login.setStyleSheet("background-color: #4CAF50; color: white;")
        self.button_test_login.clicked.connect(self.test_smtp_login)
        left_layout.addWidget(self.button_test_login)

        # 上传邮箱列表按钮
        self.button_upload_email_list = QPushButton("上传邮箱列表")
        self.button_upload_email_list.setStyleSheet("background-color: #008CBA; color: white;")
        self.button_upload_email_list.clicked.connect(self.upload_email_list)
        left_layout.addWidget(self.button_upload_email_list)

        # 邮箱列表显示
        email_list_group = QGroupBox("邮箱列表 (0)")
        email_list_layout = QVBoxLayout()
        self.listbox_email_list = QListWidget()
        email_list_layout.addWidget(self.listbox_email_list)
        email_list_group.setLayout(email_list_layout)
        left_layout.addWidget(email_list_group)
        self.email_list_group = email_list_group  # 保存引用以便更新标题

        # 上传附件按钮
        self.button_upload_attachment = QPushButton("上传附件")
        self.button_upload_attachment.setStyleSheet("background-color: #be14807f; color: black;")
        self.button_upload_attachment.clicked.connect(self.upload_attachment)
        left_layout.addWidget(self.button_upload_attachment)

        # 附件列表显示
        attachment_group = QGroupBox("附件列表")
        attachment_layout = QVBoxLayout()
        self.listbox_attachments = QListWidget()
        attachment_layout.addWidget(self.listbox_attachments)
        attachment_group.setLayout(attachment_layout)
        left_layout.addWidget(attachment_group)

        # 邮件主题和正文
        message_group = QGroupBox("邮件内容")
        message_layout = QFormLayout()
        self.entry_subject = QLineEdit()
        self.entry_subject.setPlaceholderText("请输入邮件主题")
        self.text_message = QTextEdit()
        self.text_message.setPlaceholderText("请输入邮件正文")
        message_layout.addRow("邮件主题:", self.entry_subject)
        message_layout.addRow("邮件正文:", self.text_message)
        message_group.setLayout(message_layout)
        left_layout.addWidget(message_group)

        # 发送邮件按钮
        self.button_send_emails = QPushButton("发送邮件")
        self.button_send_emails.setStyleSheet("background-color: #008CBA; color: white;")
        self.button_send_emails.clicked.connect(self.send_emails)
        left_layout.addWidget(self.button_send_emails)

        # 右侧布局
        right_widget = QWidget()
        right_layout = QVBoxLayout()
        right_widget.setLayout(right_layout)

        # 发送状态表格
        status_group = QGroupBox("发送状态")
        status_layout = QVBoxLayout()
        self.table_status = QTableWidget()
        self.table_status.setColumnCount(2)
        self.table_status.setHorizontalHeaderLabels(["邮箱", "状态"])
        self.table_status.setColumnWidth(0, 300)  # 设置邮箱列宽度
        self.table_status.setColumnWidth(1, 100)  # 设置状态列宽度
        self.table_status.horizontalHeader().setStretchLastSection(True)  # 最后一列自动拉伸
        status_layout.addWidget(self.table_status)
        status_group.setLayout(status_layout)
        right_layout.addWidget(status_group)

        # 进度条
        self.progress_bar = QProgressBar()
        right_layout.addWidget(self.progress_bar)

        # 将左右布局添加到主布局
        main_layout.addWidget(left_widget)
        main_layout.addWidget(right_widget)

        # 设置全局样式
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

    def update_email_count(self):
        """更新邮箱数量显示"""
        count = len(email_list)
        self.email_list_group.setTitle(f"邮箱列表 ({count})")

    def test_smtp_login(self):
        sender_email = self.entry_sender_email.text()
        sender_pass = self.entry_password.text()
        smtp = None
        try:
            smtp_server, smtp_port = get_smtp_server(sender_email)
            if not smtp_server:
                show_custom_message_box(QMessageBox.Critical, "不支持的邮箱域名", "错误")
                return
            smtp = smtplib.SMTP_SSL(smtp_server, smtp_port)
            smtp.login(sender_email, sender_pass)
            self.is_logged_in = True  # 登录成功，设置状态为 True
            show_custom_message_box(QMessageBox.Information, "SMTP登录成功！", "成功")
        except Exception as e:
            self.is_logged_in = False  # 登录失败，设置状态为 False
            show_custom_message_box(QMessageBox.Critical, f"SMTP登录失败: {e}", "错误")
        finally:
            if smtp:
                smtp.quit()

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
                    if cell[0] is not None and not cell[0].endswith('@yahoo.com'):  # 过滤掉 yahoo.com 邮箱
                        column_values.append(cell[0])
                email_list.extend(column_values)
            elif file_path.endswith('.txt'):
                with open(file_path, 'r', encoding='utf-8') as f:
                    email_data = f.read().split('\n')
                    for email in email_data:
                        if email.strip() and not email.strip().startswith('#') and not email.strip().endswith('@yahoo.com'):  # 过滤掉 yahoo.com 邮箱
                            email_list.append(email.strip())
            self.listbox_email_list.clear()
            for email in email_list:
                self.listbox_email_list.addItem(email)
            self.update_email_count()  # 更新邮箱数量显示
            show_custom_message_box(QMessageBox.Information, "邮箱列表上传成功！", "成功")

    def upload_attachment(self):
        file_paths, _ = QFileDialog.getOpenFileNames(self, "选择附件", "", "All files (*.*)")
        if file_paths:
            for file_path in file_paths:
                if os.path.getsize(file_path) > 10 * 1024 * 1024:  # 限制文件大小为10MB
                    show_custom_message_box(QMessageBox.Warning, f"文件 {file_path} 超过10MB，跳过此附件。", "警告")
                    continue
                attachment_paths.append(file_path)
                self.listbox_attachments.addItem(file_path)
            show_custom_message_box(QMessageBox.Information, "附件上传成功！", "成功")

    def send_emails(self):
        # 检查是否已登录
        if not self.is_logged_in:
            show_custom_message_box(QMessageBox.Warning, "请先测试SMTP登录并成功登录！", "警告")
            return

        sender_email = self.entry_sender_email.text()
        sender_pass = self.entry_password.text()
        subject = self.entry_subject.text()
        message = self.text_message.toPlainText()  # 获取纯文本内容

        if not sender_email or not sender_pass or not subject or not message:
            show_custom_message_box(QMessageBox.Warning, "请填写所有必填项！", "警告")
            return

        success_count = 0  # 成功发送的邮件数量
        failure_count = 0  # 发送失败的邮件数量
        failure_emails = []  # 存储发送失败的邮箱号

        self.table_status.setRowCount(len(email_list))
        self.progress_bar.setMaximum(len(email_list))
        self.progress_bar.setValue(0)

        for i, email in enumerate(email_list):
            success, failed_email = send_email(sender_email, sender_pass, email, subject, message, attachment_paths)
            if success:
                success_count += 1
                self.table_status.setItem(i, 0, QTableWidgetItem(email))
                self.table_status.setItem(i, 1, QTableWidgetItem("成功"))
            else:
                failure_count += 1
                failure_emails.append(failed_email)
                self.table_status.setItem(i, 0, QTableWidgetItem(email))
                self.table_status.setItem(i, 1, QTableWidgetItem("失败"))
            self.progress_bar.setValue(i + 1)
            QApplication.processEvents()  # 更新UI

        # 显示发送结果
        show_custom_message_box(QMessageBox.Information, f"成功发送 {success_count} 封邮件，失败 {failure_count} 封。", "发送结果")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = EmailSenderApp()
    ex.show()  # 确保窗口显示
    sys.exit(app.exec_())
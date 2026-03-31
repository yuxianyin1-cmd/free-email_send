import sys
import os
import time
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr
from configparser import ConfigParser
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QLineEdit, QFileDialog, QProgressBar, QTextEdit, QGroupBox,
    QMessageBox, QTabWidget, QSpinBox, QFormLayout, QFrame
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QFont, QIcon

# ===================== 全局配置管理（自动保存/加载） =====================
CONFIG_FILE = "config.ini"

def init_config():
    """初始化配置文件，不存在则创建默认配置"""
    if not os.path.exists(CONFIG_FILE):
        cfg = ConfigParser()
        cfg["sender"] = {
            "email": "",
            "auth_code": "",
            "sender_name": "人力资源部",
            "smtp_server": "",
            "smtp_port": "",
            "timeout": "20"
        }
        cfg["send"] = {
            "sleep_per_mail": "8",
            "retry_times": "1",
            "enable_log": "True"
        }
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            cfg.write(f)

def load_config():
    """加载配置文件"""
    cfg = ConfigParser()
    cfg.read(CONFIG_FILE, encoding="utf-8")
    return cfg

def save_config(section, key, value):
    """保存配置到文件"""
    cfg = ConfigParser()
    cfg.read(CONFIG_FILE, encoding="utf-8")
    if section not in cfg:
        cfg[section] = {}
    cfg[section][key] = str(value)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        cfg.write(f)

# ===================== 长连接邮件发送器（稳定防风控） =====================
class LongConnectionEmailSender:
    def __init__(self, cfg):
        self.cfg = cfg
        self.smtp = None
        self.is_connected = False

    def connect(self):
        """建立SMTP连接"""
        try:
            self.quit()
            sender_cfg = self.cfg["sender"]
            self.smtp = smtplib.SMTP_SSL(
                sender_cfg["smtp_server"],
                int(sender_cfg["smtp_port"]),
                timeout=int(sender_cfg["timeout"])
            )
            self.smtp.login(sender_cfg["email"], sender_cfg["auth_code"])
            self.is_connected = True
            return True, "连接成功"
        except smtplib.SMTPAuthenticationError:
            self.is_connected = False
            return False, "授权码错误！请检查邮箱授权码是否正确，是否开启了POP3/SMTP服务"
        except smtplib.SMTPConnectError:
            self.is_connected = False
            return False, "连接服务器失败！请检查SMTP地址和端口是否正确"
        except Exception as e:
            self.is_connected = False
            return False, f"连接异常：{str(e)}"

    def send_single(self, emp):
        """发送单封邮件，带重试"""
        sender_cfg = self.cfg["sender"]
        send_cfg = self.cfg["send"]
        last_error = "未知错误"
        max_retry = int(send_cfg["retry_times"])

        for attempt in range(max_retry + 1):
            try:
                if not self.is_connected or not self.smtp:
                    connect_ok, connect_msg = self.connect()
                    if not connect_ok:
                        last_error = connect_msg
                        time.sleep(3)
                        continue

                # 生成工资条HTML
                try:
                    emp["应发工资"] = emp["基本工资"] + emp["提成"] + emp["加班工资"] - emp["社保扣除"] - emp["考勤扣除"]
                except:
                    emp["应发工资"] = "计算异常"

                html_content = f"""
<html>
<body style="font-family:Microsoft YaHei;font-size:14px;">
    <h3>【工资条】{emp['姓名']} 您好</h3>
    <p>本月工资明细如下：</p>
    <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;width:95%;">
        <tr bgcolor="#f5f5f5" align="center">
            <th>工号</th><th>姓名</th><th>部门</th><th>基本工资</th><th>提成</th><th>加班工资</th><th>社保扣除</th><th>考勤扣除</th><th>应发工资</th>
        </tr>
        <tr align="center">
            <td>{emp['工号']}</td><td>{emp['姓名']}</td><td>{emp['部门']}</td>
            <td>{emp['基本工资']}</td><td>{emp['提成']}</td><td>{emp['加班工资']}</td>
            <td>{emp['社保扣除']}</td><td>{emp['考勤扣除']}</td>
            <td style="color:red;font-weight:bold;">{emp['应发工资']}</td>
        </tr>
    </table>
    <p style="margin-top:20px;color:#666;">系统自动发送，请勿回复</p>
</body>
</html>
                """

                msg = MIMEText(html_content, "html", "utf-8")
                msg["Subject"] = Header(f"工资条_{emp['姓名']}_{emp['工号']}", "utf-8")
                msg["From"] = formataddr((sender_cfg["sender_name"], sender_cfg["email"]), "utf-8")
                msg["To"] = str(emp["邮箱"])

                self.smtp.sendmail(sender_cfg["email"], [str(emp["邮箱"])], msg.as_string())
                return (True, emp, "")

            except smtplib.SMTPException as e:
                last_error = f"服务器拒绝：{str(e)}"
                self.is_connected = False
                if attempt < max_retry:
                    time.sleep(5)
            except Exception as e:
                last_error = f"发送错误：{str(e)}"
                self.is_connected = False
                if attempt < max_retry:
                    time.sleep(3)

        return (False, emp, last_error)

    def quit(self):
        """安全关闭连接"""
        try:
            if self.smtp:
                self.smtp.quit()
        except:
            pass
        self.is_connected = False
        self.smtp = None

# ===================== 发送线程（不卡界面） =====================
class EmailSenderThread(QThread):
    log_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int, int, int)
    finished_signal = pyqtSignal(int, int, list)

    def __init__(self, cfg, excel_path):
        super().__init__()
        self.cfg = cfg
        self.excel_path = excel_path
        self._is_running = True
        self.sent_log_file = os.path.join(os.path.dirname(excel_path), "已发送名单.txt")
        self.sender = LongConnectionEmailSender(cfg)

    def load_sent_log(self):
        """加载已发送名单"""
        if self.cfg["send"].getboolean("enable_log") and os.path.exists(self.sent_log_file):
            try:
                with open(self.sent_log_file, "r", encoding="utf-8") as f:
                    return set(line.strip() for line in f if line.strip())
            except:
                return set()
        return set()

    def save_sent_log(self, emp_id):
        """记录已发送"""
        if self.cfg["send"].getboolean("enable_log"):
            try:
                with open(self.sent_log_file, "a", encoding="utf-8") as f:
                    f.write(f"{emp_id}\n")
            except:
                pass

    def run(self):
        try:
            # 读取Excel
            self.log_signal.emit("📂 正在读取员工工资表...")
            try:
                df = pd.read_excel(self.excel_path, engine="openpyxl")
                df = df.dropna(how="all")
                all_employees = df.to_dict("records")
            except Exception as e:
                self.log_signal.emit(f"❌ 读取Excel失败：{str(e)}")
                self.log_signal.emit("⚠️  请检查Excel文件是否关闭，列名是否符合模板要求")
                self.finished_signal.emit(0, 0, [])
                return

            total_all = len(all_employees)
            if total_all == 0:
                self.log_signal.emit("❌ Excel中没有有效员工数据！")
                self.finished_signal.emit(0, 0, [])
                return

            # 过滤已发送
            sent_ids = self.load_sent_log()
            employees = [emp for emp in all_employees if str(emp.get("工号", "")) not in sent_ids]
            total_remaining = len(employees)
            total_sent = total_all - total_remaining

            self.log_signal.emit(f"✅ 总人数：{total_all} | 已发送：{total_sent} | 待发送：{total_remaining}")
            if total_remaining == 0:
                self.log_signal.emit("🎉 所有员工工资条均已发送完成！")
                self.finished_signal.emit(total_sent, 0, [])
                return

            # 连接服务器
            self.log_signal.emit("🔗 正在连接邮件服务器...")
            connect_ok, connect_msg = self.sender.connect()
            if not connect_ok:
                self.log_signal.emit(f"❌ {connect_msg}")
                self.finished_signal.emit(0, total_remaining, employees)
                return
            self.log_signal.emit("✅ 邮件服务器连接成功，开始发送...")

            sleep_per_mail = int(self.cfg["send"]["sleep_per_mail"])
            self.log_signal.emit(f"⚠️  发送间隔：{sleep_per_mail}秒/封，触发频率限制会自动等待1小时")

            # 循环发送
            success = 0
            failed_list = []
            idx = 0
            while idx < len(employees):
                emp = employees[idx]
                if not self._is_running:
                    self.log_signal.emit("⚠️  用户手动停止了发送任务")
                    break

                is_success, emp, error = self.sender.send_single(emp)

                if is_success:
                    success += 1
                    self.save_sent_log(str(emp.get("工号", "")))
                    self.log_signal.emit(f"✅ [{idx+1}/{total_remaining}] {emp.get('姓名', '未知')} 发送成功")
                    idx += 1
                    if idx < len(employees) and self._is_running:
                        time.sleep(sleep_per_mail)
                else:
                    # 处理550频率限制
                    if "550" in error and "Too many attempts" in error:
                        self.log_signal.emit(f"⚠️  触发邮箱频率限制！自动等待60分钟后继续...")
                        # 分段等待，支持中途停止
                        for wait_min in range(60):
                            if not self._is_running:
                                break
                            time.sleep(60)
                            self.log_signal.emit(f"⏳ 已等待 {wait_min+1}/60 分钟...")
                        # 等待结束重连
                        self.log_signal.emit(f"🔄 等待结束，重新连接服务器...")
                        self.sender.quit()
                        time.sleep(5)
                        self.sender.connect()
                    else:
                        # 其他错误记录失败
                        failed_list.append({"emp": emp, "error": error})
                        self.log_signal.emit(f"❌ [{idx+1}/{total_remaining}] {emp.get('姓名', '未知')} 发送失败：{error}")
                        idx += 1
                        if idx < len(employees) and self._is_running:
                            time.sleep(sleep_per_mail * 2)

                # 更新进度
                self.progress_signal.emit(idx, success, len(failed_list))

            # 发送完成
            self.sender.quit()
            self.log_signal.emit(f"\n{'='*50}")
            self.log_signal.emit(f"🎉 发送任务结束！")
            self.log_signal.emit(f"   成功发送：{success} 封 | 发送失败：{len(failed_list)} 封")
            self.log_signal.emit(f"{'='*50}")
            self.finished_signal.emit(success, len(failed_list), failed_list)

        except Exception as e:
            self.log_signal.emit(f"❌ 程序异常：{str(e)}")
            import traceback
            self.log_signal.emit(traceback.format_exc())
            self.sender.quit()

    def stop(self):
        self._is_running = False
        self.sender.quit()

# ===================== 主界面 =====================
class MainWindow(QTabWidget):
    def __init__(self):
        super().__init__()
        init_config()
        self.cfg = load_config()
        self.sender_thread = None
        self.failed_list = []
        self.initUI()

    def initUI(self):
        self.setWindowTitle("🚀 自动化工资条邮件发送工具")
        self.setGeometry(200, 200, 900, 700)
        self.setFont(QFont("Microsoft YaHei", 10))
        self.setStyleSheet("""
            QWidget { font-family: 'Microsoft YaHei'; font-size: 14px; }
            QPushButton { padding: 8px 16px; background-color: #0078d4; color: white; border-radius: 4px; min-height: 20px; }
            QPushButton:hover { background-color: #106ebe; }
            QPushButton:disabled { background-color: #cccccc; }
            QPushButton#danger { background-color: #d13438; }
            QPushButton#danger:hover { background-color: #a4262c; }
            QPushButton#success { background-color: #107c10; }
            QPushButton#success:hover { background-color: #0b5c0b; }
            QLineEdit, QSpinBox { padding: 6px; border: 1px solid #ddd; border-radius: 4px; }
            QTextEdit { border: 1px solid #ddd; border-radius: 4px; background-color: #fafafa; }
            QGroupBox { font-weight: bold; margin-top: 10px; padding-top: 15px; }
            QTabWidget::pane { border: 1px solid #ddd; border-radius: 4px; }
            QTabBar::tab { padding: 8px 20px; border-radius: 4px 4px 0 0; margin-right: 2px; }
            QTabBar::tab:selected { background-color: #0078d4; color: white; }
        """)

        # 1. 发送主页面
        self.send_tab = QWidget()
        self.init_send_tab()
        self.addTab(self.send_tab, "📤 发送工资条")

        # 2. 配置页面
        self.config_tab = QWidget()
        self.init_config_tab()
        self.addTab(self.config_tab, "⚙️ 邮箱配置")

        # 3. 帮助页面
        self.help_tab = QWidget()
        self.init_help_tab()
        self.addTab(self.help_tab, "📖 使用帮助")

    # ===================== 发送页面初始化 =====================
    def init_send_tab(self):
        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        # 1. 文件选择区域
        file_group = QGroupBox("📂 选择员工工资表")
        file_layout = QHBoxLayout()

        self.excel_path_edit = QLineEdit()
        self.excel_path_edit.setPlaceholderText("请选择Excel格式的员工工资表，点击右侧按钮生成模板")
        file_layout.addWidget(self.excel_path_edit, stretch=5)

        self.select_file_btn = QPushButton("选择文件")
        self.select_file_btn.clicked.connect(self.select_excel_file)
        file_layout.addWidget(self.select_file_btn)

        self.create_template_btn = QPushButton("生成Excel模板")
        self.create_template_btn.setObjectName("success")
        self.create_template_btn.clicked.connect(self.create_excel_template)
        file_layout.addWidget(self.create_template_btn)

        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # 2. 进度区域
        progress_group = QGroupBox("📊 发送进度")
        progress_layout = QVBoxLayout()

        self.progress_bar = QProgressBar()
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.progress_bar.setFormat("%p%")
        progress_layout.addWidget(self.progress_bar)

        self.status_label = QLabel("等待开始发送...")
        self.status_label.setStyleSheet("color: #555;")
        progress_layout.addWidget(self.status_label)

        progress_group.setLayout(progress_layout)
        layout.addWidget(progress_group)

        # 3. 日志区域
        log_group = QGroupBox("📝 发送日志")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 9))
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group, stretch=2)

        # 4. 按钮区域
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        self.export_failed_btn = QPushButton("导出失败名单")
        self.export_failed_btn.clicked.connect(self.export_failed_list)
        self.export_failed_btn.setEnabled(False)
        btn_layout.addWidget(self.export_failed_btn)

        self.start_btn = QPushButton("▶️  开始发送")
        self.start_btn.setFixedSize(140, 45)
        self.start_btn.clicked.connect(self.start_sending)
        btn_layout.addWidget(self.start_btn)

        self.stop_btn = QPushButton("⏹️  停止发送")
        self.stop_btn.setObjectName("danger")
        self.stop_btn.setFixedSize(140, 45)
        self.stop_btn.clicked.connect(self.stop_sending)
        self.stop_btn.setEnabled(False)
        btn_layout.addWidget(self.stop_btn)

        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        self.send_tab.setLayout(layout)

    # ===================== 配置页面初始化 =====================
    def init_config_tab(self):
        layout = QVBoxLayout()
        layout.setSpacing(20)
        layout.setContentsMargins(20, 20, 20, 20)

        # 发件人配置
        sender_group = QGroupBox("📧 发件人邮箱配置")
        form_layout = QFormLayout()
        form_layout.setSpacing(15)
        form_layout.setLabelAlignment(Qt.AlignRight)

        self.sender_email_edit = QLineEdit(self.cfg["sender"].get("email", ""))
        form_layout.addRow("发件人邮箱：", self.sender_email_edit)

        self.auth_code_edit = QLineEdit(self.cfg["sender"].get("auth_code", ""))
        self.auth_code_edit.setEchoMode(QLineEdit.Password)
        form_layout.addRow("邮箱授权码：", self.auth_code_edit)

        self.sender_name_edit = QLineEdit(self.cfg["sender"].get("sender_name", "人力资源部"))
        form_layout.addRow("发件人名称：", self.sender_name_edit)

        self.smtp_server_edit = QLineEdit(self.cfg["sender"].get("smtp_server", "smtp.qq.com"))
        form_layout.addRow("SMTP服务器地址：", self.smtp_server_edit)

        self.smtp_port_edit = QLineEdit(self.cfg["sender"].get("smtp_port", "465"))
        form_layout.addRow("SMTP端口：", self.smtp_port_edit)

        # 测试连接按钮
        test_btn_layout = QHBoxLayout()
        test_btn_layout.addStretch()
        self.test_connect_btn = QPushButton("🔍 测试邮箱连接")
        self.test_connect_btn.clicked.connect(self.test_smtp_connect)
        test_btn_layout.addWidget(self.test_connect_btn)
        form_layout.addRow("", test_btn_layout)

        sender_group.setLayout(form_layout)
        layout.addWidget(sender_group)

        # 发送策略配置
        send_group = QGroupBox("⚙️ 发送策略配置")
        send_form_layout = QFormLayout()
        send_form_layout.setSpacing(15)
        send_form_layout.setLabelAlignment(Qt.AlignRight)

        self.sleep_spin = QSpinBox()
        self.sleep_spin.setRange(1, 60)
        self.sleep_spin.setValue(int(self.cfg["send"].get("sleep_per_mail", "8")))
        send_form_layout.addRow("每封邮件间隔(秒)：", self.sleep_spin)

        self.retry_spin = QSpinBox()
        self.retry_spin.setRange(0, 5)
        self.retry_spin.setValue(int(self.cfg["send"].get("retry_times", "1")))
        send_form_layout.addRow("失败重试次数：", self.retry_spin)

        # 保存配置按钮
        save_btn_layout = QHBoxLayout()
        save_btn_layout.addStretch()
        self.save_config_btn = QPushButton("💾 保存配置")
        self.save_config_btn.setObjectName("success")
        self.save_config_btn.clicked.connect(self.save_all_config)
        save_btn_layout.addWidget(self.save_config_btn)
        send_form_layout.addRow("", save_btn_layout)

        send_group.setLayout(send_form_layout)
        layout.addWidget(send_group)

        layout.addStretch()
        self.config_tab.setLayout(layout)

    # ===================== 帮助页面初始化 =====================
    def init_help_tab(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(30, 30, 30, 30)

        help_text = QTextEdit()
        help_text.setReadOnly(True)
        help_text.setFont(QFont("Microsoft YaHei", 12))
        help_text.setHtml("""
<h2>📖 自动化工资条邮件发送工具 使用说明</h2>
<hr>
<h3>一、首次使用步骤</h3>
<p>1. 切换到【⚙️ 邮箱配置】页面，填写你的发件人邮箱、授权码、SMTP信息</p>
<p>2. 点击【测试邮箱连接】，提示连接成功后，点击【保存配置】</p>
<p>3. 回到【📤 发送工资条】页面，点击【生成Excel模板】，在桌面生成标准工资表</p>
<p>4. 打开模板，填写员工工资信息，保存关闭Excel</p>
<p>5. 点击【选择文件】，选择填好的工资表Excel</p>
<p>6. 点击【开始发送】，等待发送完成即可</p>

<hr>
<h3>二、常见问题</h3>
<p><b>1. QQ邮箱授权码怎么获取？</b></p>
<p>答：登录QQ邮箱网页版 → 【设置】→ 【账户】→ 开启【POP3/SMTP服务】→ 按提示生成授权码</p>

<p><b>2. 提示"连接服务器失败"怎么办？</b></p>
<p>答：检查SMTP地址和端口是否正确，QQ邮箱默认smtp.qq.com:465，企业邮箱请咨询公司IT</p>

<p><b>3. 提示"授权码错误"怎么办？</b></p>
<p>答：确认开启了POP3/SMTP服务，重新生成授权码，复制时不要带空格</p>

<p><b>4. 发送失败提示"Too many attempts"怎么办？</b></p>
<p>答：这是个人邮箱的频率限制，程序会自动等待1小时后继续发送，也可以换用企业邮箱</p>

<p><b>5. 中途停止后，下次会重复发送吗？</b></p>
<p>答：不会，程序会自动记录已发送的员工工号，下次打开会自动跳过已发送的人</p>

<hr>
<h3>三、Excel模板必填列</h3>
<p>Excel必须包含以下列名（和模板完全一致，不要改字）：</p>
<p><b>工号、姓名、部门、基本工资、提成、加班工资、社保扣除、考勤扣除、邮箱</b></p>
<p>其他列可以自行添加，不影响程序读取</p>
        """)
        layout.addWidget(help_text)
        self.help_tab.setLayout(layout)

    # ===================== 功能函数 =====================
    def append_log(self, message):
        self.log_text.append(message)
        self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())

    def update_progress(self, current, success, failed):
        total = self.total_remaining
        self.progress_bar.setValue(int((current / total) * 100))
        self.status_label.setText(f"当前：{current}/{total} | 成功：{success} | 失败：{failed}")

    def select_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择员工工资表", "", "Excel文件 (*.xlsx *.xls)"
        )
        if file_path:
            self.excel_path_edit.setText(file_path)
            self.append_log(f"✅ 已选择工资表：{os.path.basename(file_path)}")

    def create_excel_template(self):
        """生成Excel模板到桌面"""
        try:
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            template_path = os.path.join(desktop_path, "工资条模板.xlsx")

            # 生成模板数据
            template_data = {
                "工号": ["10001", "10002"],
                "姓名": ["张三", "李四"],
                "部门": ["技术部", "人事部"],
                "基本工资": [8000, 7000],
                "提成": [2000, 1000],
                "加班工资": [500, 0],
                "社保扣除": [800, 700],
                "考勤扣除": [0, 100],
                "邮箱": ["zhangsan@example.com", "lisi@example.com"]
            }
            df = pd.DataFrame(template_data)
            df.to_excel(template_path, index=False, engine="openpyxl")

            QMessageBox.information(self, "生成成功", f"Excel模板已生成到桌面！\n路径：{template_path}")
            self.append_log("✅ 已生成Excel模板到桌面")
        except Exception as e:
            QMessageBox.warning(self, "生成失败", f"模板生成失败：{str(e)}")

    def test_smtp_connect(self):
        """测试SMTP连接"""
        # 先保存当前填写的配置
        temp_cfg = ConfigParser()
        temp_cfg["sender"] = {
            "email": self.sender_email_edit.text().strip(),
            "auth_code": self.auth_code_edit.text().strip(),
            "sender_name": self.sender_name_edit.text().strip(),
            "smtp_server": self.smtp_server_edit.text().strip(),
            "smtp_port": self.smtp_port_edit.text().strip(),
            "timeout": "20"
        }
        temp_cfg["send"] = self.cfg["send"]

        sender = LongConnectionEmailSender(temp_cfg)
        connect_ok, connect_msg = sender.connect()
        sender.quit()

        if connect_ok:
            QMessageBox.information(self, "测试成功", "✅ 邮箱连接测试成功！配置正确")
            self.append_log("✅ 邮箱连接测试成功")
        else:
            QMessageBox.warning(self, "测试失败", connect_msg)
            self.append_log(f"❌ 邮箱连接测试失败：{connect_msg}")

    def save_all_config(self):
        """保存所有配置"""
        save_config("sender", "email", self.sender_email_edit.text().strip())
        save_config("sender", "auth_code", self.auth_code_edit.text().strip())
        save_config("sender", "sender_name", self.sender_name_edit.text().strip())
        save_config("sender", "smtp_server", self.smtp_server_edit.text().strip())
        save_config("sender", "smtp_port", self.smtp_port_edit.text().strip())
        save_config("send", "sleep_per_mail", self.sleep_spin.value())
        save_config("send", "retry_times", self.retry_spin.value())

        # 重新加载配置
        self.cfg = load_config()
        QMessageBox.information(self, "保存成功", "✅ 配置已保存！")
        self.append_log("✅ 配置已保存")

    def start_sending(self):
        """开始发送"""
        excel_path = self.excel_path_edit.text().strip()
        if not excel_path or not os.path.exists(excel_path):
            QMessageBox.warning(self, "提示", "请先选择有效的员工工资表Excel文件！")
            return

        # 校验配置
        sender_cfg = self.cfg["sender"]
        if not sender_cfg.get("email") or not sender_cfg.get("auth_code"):
            QMessageBox.warning(self, "提示", "请先在【邮箱配置】页面填写发件人邮箱和授权码！")
            self.setCurrentIndex(1)
            return

        # 预读取数据
        try:
            df = pd.read_excel(excel_path, engine="openpyxl")
            required_cols = ["工号", "姓名", "邮箱"]
            for col in required_cols:
                if col not in df.columns:
                    QMessageBox.warning(self, "格式错误", f"Excel缺少必填列：{col}\n请使用【生成Excel模板】功能生成标准模板")
                    return

            sent_ids = set()
            if self.cfg["send"].getboolean("enable_log"):
                sent_log_file = os.path.join(os.path.dirname(excel_path), "已发送名单.txt")
                if os.path.exists(sent_log_file):
                    with open(sent_log_file, "r", encoding="utf-8") as f:
                        sent_ids = set(line.strip() for line in f if line.strip())
            self.total_remaining = len([emp for emp in df.to_dict("records") if str(emp.get("工号", "")) not in sent_ids])
        except Exception as e:
            QMessageBox.warning(self, "文件错误", f"读取工资表失败：{str(e)}\n请检查Excel文件是否关闭")
            return

        if self.total_remaining == 0:
            QMessageBox.information(self, "提示", "该工资表内所有员工均已发送完成！")
            return

        # 重置界面
        self.log_text.clear()
        self.progress_bar.setValue(0)
        self.failed_list = []
        self.export_failed_btn.setEnabled(False)

        # 启动线程
        self.sender_thread = EmailSenderThread(self.cfg, excel_path)
        self.sender_thread.log_signal.connect(self.append_log)
        self.sender_thread.progress_signal.connect(self.update_progress)
        self.sender_thread.finished_signal.connect(self.on_send_finished)
        self.sender_thread.start()

        # 更新按钮状态
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.select_file_btn.setEnabled(False)
        self.create_template_btn.setEnabled(False)
        self.excel_path_edit.setEnabled(False)

    def stop_sending(self):
        """停止发送"""
        if self.sender_thread:
            self.sender_thread.stop()
            self.append_log("⏹️  正在停止发送任务...")
            self.stop_btn.setEnabled(False)

    def on_send_finished(self, success, failed, failed_list):
        """发送完成回调"""
        self.failed_list = failed_list
        # 恢复按钮状态
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.select_file_btn.setEnabled(True)
        self.create_template_btn.setEnabled(True)
        self.excel_path_edit.setEnabled(True)
        self.export_failed_btn.setEnabled(len(failed_list) > 0)

        # 弹出提示
        QMessageBox.information(
            self, "任务完成",
            f"工资条发送任务已结束！\n\n✅ 成功发送：{success} 封\n❌ 发送失败：{failed} 封"
        )

    def export_failed_list(self):
        """导出失败名单"""
        if not self.failed_list:
            QMessageBox.warning(self, "提示", "没有发送失败的记录！")
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self, "保存失败名单", os.path.join(os.path.expanduser("~"), "Desktop/发送失败名单.xlsx"),
            "Excel文件 (*.xlsx)"
        )
        if save_path:
            try:
                export_data = []
                for item in self.failed_list:
                    emp = item["emp"]
                    export_data.append({
                        "工号": emp.get("工号", ""),
                        "姓名": emp.get("姓名", ""),
                        "部门": emp.get("部门", ""),
                        "邮箱": emp.get("邮箱", ""),
                        "失败原因": item["error"]
                    })
                df = pd.DataFrame(export_data)
                df.to_excel(save_path, index=False, engine="openpyxl")
                QMessageBox.information(self, "导出成功", f"失败名单已导出到：{save_path}")
            except Exception as e:
                QMessageBox.warning(self, "导出失败", f"导出失败：{str(e)}")

# ===================== 程序入口 =====================
if __name__ == '__main__':
    # 解决Windows高DPI缩放问题
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
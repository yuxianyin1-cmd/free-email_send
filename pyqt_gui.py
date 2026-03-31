import sys
import os
import time
import traceback
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QLineEdit, QFileDialog, QProgressBar, QTextEdit, QGroupBox, QMessageBox
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QFont

# 导入模块
from config import (
    SENDER_EMAIL, SENDER_NAME, SMTP_SERVER, SMTP_PORT,
    MAX_WORKERS, RETRY_TIMES, SLEEP_PER_MAIL, ENABLE_LOG
)
from excel_utils import read_employee_data
from email_utils import LongConnectionEmailSender


# ===================== 发送邮件的工作线程（长连接） =====================
class EmailSenderThread(QThread):
    # 界面更新信号
    log_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int, int, int)
    finished_signal = pyqtSignal(int, int, list)

    def __init__(self, excel_path):
        super().__init__()
        self.excel_path = excel_path
        self._is_running = True
        self.sent_log_file = os.path.join(os.path.dirname(excel_path), "已发送名单.txt")
        self.sender = LongConnectionEmailSender()  # 初始化长连接发送器

    def load_sent_log(self):
        """加载已发送名单（断点续传）"""
        if not ENABLE_LOG or not os.path.exists(self.sent_log_file):
            return set()
        try:
            with open(self.sent_log_file, "r", encoding="utf-8") as f:
                return set(line.strip() for line in f if line.strip())
        except:
            return set()

    def save_sent_log(self, emp_id):
        """记录已发送员工工号"""
        if not ENABLE_LOG:
            return
        try:
            with open(self.sent_log_file, "a", encoding="utf-8") as f:
                f.write(f"{emp_id}\n")
        except:
            pass

    def run(self):
        try:
            # 1. 读取Excel数据
            self.log_signal.emit("📂 正在读取员工工资表...")
            all_employees = read_employee_data()
            total_all = len(all_employees)
            if total_all == 0:
                self.log_signal.emit("❌ Excel文件中未读取到有效员工数据！")
                self.finished_signal.emit(0, 0, [])
                return

            # 2. 过滤已发送人员，断点续传
            sent_ids = self.load_sent_log()
            employees = [emp for emp in all_employees if str(emp.get("工号", "")) not in sent_ids]
            total_remaining = len(employees)
            total_sent = total_all - total_remaining

            self.log_signal.emit(f"✅ 总人数：{total_all} | 已发送：{total_sent} | 待发送：{total_remaining}")
            if total_remaining == 0:
                self.log_signal.emit("🎉 所有员工工资条均已发送完成！")
                self.finished_signal.emit(total_sent, 0, [])
                return

            # 3. 初始化长连接
            self.log_signal.emit("🔗 正在连接邮件服务器...")
            connect_result = self.sender.connect()
            if connect_result is not True:
                self.log_signal.emit(f"❌ {connect_result[1]}")
                self.finished_signal.emit(0, total_remaining, employees)
                return
            self.log_signal.emit("✅ 邮件服务器已经连接成功，开始发送...")
            self.log_signal.emit(f"⚠️  发送间隔：{SLEEP_PER_MAIL}秒/封，遇到限制自动等待1小时")
            
            # 4. 循环发送（核心：遇到550错误自动长等待）
            success = 0
            failed_list = []
            idx = 0
            while idx < len(employees):
                emp = employees[idx]

                # 检测用户停止操作
                if not self._is_running:
                    self.log_signal.emit("⚠️  用户手动停止了发送任务")
                    break

                # 发送单封邮件
                is_success, emp, error = self.sender.send_single(emp)

                # 处理结果
                if is_success:
                    success += 1
                    self.save_sent_log(str(emp.get("工号", "")))
                    self.log_signal.emit(f"✅ [{idx + 1}/{total_remaining}] {emp.get('姓名', '未知')} 发送成功")
                    idx += 1  # 成功才进入下一封
                    # 正常发送间隔
                    if idx < len(employees) and self._is_running:
                        time.sleep(SLEEP_PER_MAIL)
                else:
                    # 检查是不是550频率限制错误
                    if "550" in error and "Too many attempts" in error:
                        self.log_signal.emit(f"⚠️  即将触发QQ邮箱频率限制！自动等待60分钟后继续...")
                        # 分段等待60分钟，方便用户中途停止
                        for wait_min in range(60):
                            if not self._is_running:
                                break
                            time.sleep(60)  # 等1分钟
                            self.log_signal.emit(f"⏳ 已等待 {wait_min + 1}/60 分钟...")
                        # 等待结束后，重连服务器，重试这一封（不idx+1）
                        self.log_signal.emit(f"🔄 等待结束，重新连接服务器并尝试发送...")
                        self.sender.quit()
                        time.sleep(5)
                        self.sender.connect()
                    else:
                        # 其他错误，记录失败，进入下一封
                        failed_list.append({"emp": emp, "error": error})
                        self.log_signal.emit(
                            f"❌ [{idx + 1}/{total_remaining}] {emp.get('姓名', '未知')} 发送失败：{error}")
                        idx += 1
                        # 失败后也稍微等一下
                        if idx < len(employees) and self._is_running:
                            time.sleep(SLEEP_PER_MAIL * 2)

                # 更新进度
                self.progress_signal.emit(idx, success, len(failed_list))

            # 5. 发送完成，安全关闭连接
            self.sender.quit()
            self.log_signal.emit(f"\n{'=' * 50}")
            self.log_signal.emit(f"🎉 发送任务结束！")
            self.log_signal.emit(f"   成功发送：{success} 封 | 发送失败：{len(failed_list)} 封")
            self.log_signal.emit(f"{'=' * 50}")
            self.finished_signal.emit(success, len(failed_list), failed_list)

        except Exception as e:
            self.log_signal.emit(f"❌ 程序异常：{str(e)}")
            self.log_signal.emit(traceback.format_exc())
            self.sender.quit()


# ===================== 主界面 =====================
class EmailSenderWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.sender_thread = None
        self.initUI()

    def initUI(self):
        self.setWindowTitle("🚀 自动化工资条邮件发送系统")
        self.setGeometry(200, 200, 850, 650)
        self.setStyleSheet("""
            QWidget { font-family: 'Microsoft YaHei'; font-size: 14px; }
            QPushButton { padding: 8px 16px; background-color: #0078d4; color: white; border-radius: 4px; }
            QPushButton:hover { background-color: #106ebe; }
            QPushButton:disabled { background-color: #cccccc; }
            QPushButton#stopBtn { background-color: #d13438; }
            QPushButton#stopBtn:hover { background-color: #a4262c; }
            QLineEdit { padding: 6px; border: 1px solid #ddd; border-radius: 4px; }
            QTextEdit { border: 1px solid #ddd; border-radius: 4px; background-color: #fafafa; }
            QGroupBox { font-weight: bold; margin-top: 10px; padding-top: 10px; }
        """)

        # 主布局
        main_layout = QVBoxLayout()
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # 1. 配置信息区域
        config_group = QGroupBox("📋 基础配置")
        config_layout = QVBoxLayout()

        # Excel文件选择行
        file_layout = QHBoxLayout()
        file_layout.addWidget(QLabel("员工工资表："))
        self.excel_path_edit = QLineEdit()
        self.excel_path_edit.setPlaceholderText("请选择Excel格式的员工工资表")
        file_layout.addWidget(self.excel_path_edit)
        self.select_file_btn = QPushButton("📂 选择文件")
        self.select_file_btn.clicked.connect(self.select_excel_file)
        file_layout.addWidget(self.select_file_btn)
        config_layout.addLayout(file_layout)

        # 发件人信息展示
        info_layout = QHBoxLayout()
        info_layout.addWidget(QLabel(f"发件人：{SENDER_NAME} <{SENDER_EMAIL}>"))
        info_layout.addWidget(QLabel(f"服务器：{SMTP_SERVER}:{SMTP_PORT}"))
        info_layout.addStretch()
        config_layout.addLayout(info_layout)
        config_group.setLayout(config_layout)
        main_layout.addWidget(config_group)

        # 2. 发送进度区域
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
        main_layout.addWidget(progress_group)

        # 3. 发送日志区域
        log_group = QGroupBox("📝 发送日志")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 9))
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        main_layout.addWidget(log_group)

        # 4. 操作按钮区域
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        self.start_btn = QPushButton("▶️  开始发送")
        self.start_btn.clicked.connect(self.start_sending)
        self.start_btn.setFixedSize(120, 40)
        btn_layout.addWidget(self.start_btn)

        self.stop_btn = QPushButton("⏹️  停止发送")
        self.stop_btn.setObjectName("stopBtn")
        self.stop_btn.clicked.connect(self.stop_sending)
        self.stop_btn.setEnabled(False)
        self.stop_btn.setFixedSize(120, 40)
        btn_layout.addWidget(self.stop_btn)

        btn_layout.addStretch()
        main_layout.addLayout(btn_layout)

        self.setLayout(main_layout)

    def select_excel_file(self):
        """选择Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择员工工资表", "", "Excel文件 (*.xlsx *.xls)"
        )
        if file_path:
            self.excel_path_edit.setText(file_path)
            self.append_log(f"✅ 已选择工资表：{os.path.basename(file_path)}")

    def append_log(self, message):
        """追加日志并自动滚动到底部"""
        self.log_text.append(message)
        self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())

    def update_progress(self, current, success, failed):
        """更新进度条和状态"""
        total = self.total_remaining
        self.progress_bar.setValue(int((current / total) * 100))
        self.status_label.setText(f"当前：{current}/{total} | 成功：{success} | 失败：{failed}")

    def start_sending(self):
        """开始发送按钮逻辑"""
        excel_path = self.excel_path_edit.text()
        # 校验文件
        if not excel_path or not os.path.exists(excel_path):
            QMessageBox.warning(self, "操作提示", "请先选择有效的员工工资表Excel文件！")
            return

        # 预读取数据校验
        try:
            all_employees = read_employee_data()
            sent_ids = set()
            if ENABLE_LOG:
                sent_log_file = os.path.join(os.path.dirname(excel_path), "已发送名单.txt")
                if os.path.exists(sent_log_file):
                    with open(sent_log_file, "r", encoding="utf-8") as f:
                        sent_ids = set(line.strip() for line in f if line.strip())
            self.total_remaining = len([emp for emp in all_employees if str(emp.get("工号", "")) not in sent_ids])
        except Exception as e:
            QMessageBox.warning(self, "文件错误", f"读取工资表失败：{str(e)}")
            return

        if self.total_remaining == 0:
            QMessageBox.information(self, "操作提示", "该工资表内所有员工均已发送完成！")
            return

        # 重置界面
        self.log_text.clear()
        self.progress_bar.setValue(0)

        # 启动发送线程
        self.sender_thread = EmailSenderThread(excel_path)
        self.sender_thread.log_signal.connect(self.append_log)
        self.sender_thread.progress_signal.connect(self.update_progress)
        self.sender_thread.finished_signal.connect(self.on_send_finished)
        self.sender_thread.start()

        # 更新按钮状态
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.select_file_btn.setEnabled(False)
        self.excel_path_edit.setEnabled(False)

    def stop_sending(self):
        """停止发送按钮逻辑"""
        if self.sender_thread:
            self.sender_thread.stop()
            self.append_log("⏹️  正在停止发送任务...")
            self.stop_btn.setEnabled(False)

    def on_send_finished(self, success, failed, failed_list):
        """发送完成回调"""
        # 恢复按钮状态
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.select_file_btn.setEnabled(True)
        self.excel_path_edit.setEnabled(True)

        # 弹出完成提示
        QMessageBox.information(
            self, "任务完成",
            f"工资条发送任务已结束！\n\n✅ 成功发送：{success} 封\n❌ 发送失败：{failed} 封"
        )


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = EmailSenderWindow()
    window.show()
    sys.exit(app.exec_())
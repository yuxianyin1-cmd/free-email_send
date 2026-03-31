import smtplib
import time
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr
from config import (
    SENDER_EMAIL, SENDER_AUTH_CODE, SENDER_NAME,
    SMTP_SERVER, SMTP_PORT, SMTP_TIMEOUT, RETRY_TIMES
)

def build_salary_html(emp):
    """生成工资条HTML（保持不变，兼容原有逻辑）"""
    try:
        emp["应发工资"] = emp["基本工资"] + emp["提成"] + emp["加班工资"] - emp["社保扣除"] - emp["考勤扣除"]
    except:
        emp["应发工资"] = "计算异常"
    return f"""
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

class LongConnectionEmailSender:
    """
    长连接复用邮件发送器（核心优化）
    全程只建立1次SMTP连接，避免频繁新建连接触发风控
    支持连接断开自动重连
    """
    def __init__(self):
        self.smtp = None
        self.is_connected = False

    def connect(self):
        """建立SMTP连接（全程只调用1次，或断开后重连）"""
        try:
            # 先关闭旧连接
            self.quit()
            # 新建SSL连接
            self.smtp = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, timeout=SMTP_TIMEOUT)
            self.smtp.login(SENDER_EMAIL, SENDER_AUTH_CODE)
            self.is_connected = True
            return True
        except Exception as e:
            self.is_connected = False
            return False, f"连接失败：{str(e)}"

    def send_single(self, emp):
        """复用已有连接发送单封邮件，带自动重连"""
        last_error = "未知错误"
        for attempt in range(RETRY_TIMES + 1):
            try:
                # 检查连接状态，断开则自动重连
                if not self.is_connected or not self.smtp:
                    connect_result = self.connect()
                    if connect_result is not True:
                        last_error = connect_result[1]
                        time.sleep(3)
                        continue

                # 生成邮件内容
                html_content = build_salary_html(emp)
                msg = MIMEText(html_content, "html", "utf-8")
                msg["Subject"] = Header(f"工资条_{emp['姓名']}_{emp['工号']}", "utf-8")
                msg["From"] = formataddr((SENDER_NAME, SENDER_EMAIL), "utf-8")
                msg["To"] = str(emp["邮箱"])

                # 复用连接发送
                self.smtp.sendmail(SENDER_EMAIL, [str(emp["邮箱"])], msg.as_string())
                return (True, emp, "")

            except smtplib.SMTPException as e:
                last_error = f"服务器拒绝：{str(e)}"
                self.is_connected = False  # 标记连接失效
                if attempt < RETRY_TIMES:
                    time.sleep(5)  # 失败后等待久一点再重试
            except Exception as e:
                last_error = f"发送错误：{str(e)}"
                self.is_connected = False
                if attempt < RETRY_TIMES:
                    time.sleep(3)

        # 所有重试都失败
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
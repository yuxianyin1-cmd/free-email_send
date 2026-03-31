import pandas as pd
from config import EXCEL_FILE

def read_employee_data():
    """
    从Excel读取员工数据
    :return: 员工字典列表
    """
    try:
        df = pd.read_excel(EXCEL_FILE, engine='openpyxl')
        # 去除空行
        df = df.dropna(how='all')
        return df.to_dict("records")
    except FileNotFoundError:
        print(f"❌ 找不到文件：{EXCEL_FILE}")
        print("   请检查 config.py 里的 EXCEL_FILE 路径是否正确")
        exit()
    except Exception as e:
        print(f"❌ 读取Excel失败：{e}")
        exit()
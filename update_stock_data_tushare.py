import akshare as ak
import openpyxl
import pandas as pd
from datetime import datetime, timedelta
import os
import argparse
import shutil
import time
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# 设置环境变量，避免代理问题
os.environ["NO_PROXY"] = "push2his.eastmoney.com,33.push2his.eastmoney.com,push2.eastmoney.com,88.push2.eastmoney.com"

def get_retry_session():
    s = requests.Session()
    s.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36"
        ),
        "Referer": "https://quote.eastmoney.com/",
        "Accept": "application/json, text/plain, */*",
    })
    retries = Retry(total=3, backoff_factor=1.0, allowed_methods=["GET"])
    s.mount("https://", HTTPAdapter(max_retries=retries))
    return s

session = get_retry_session()
try:
    ak.session = session
except Exception:
    pass

def create_backup(file_name):
    """
    文件备份
    """
    if not os.path.exists(file_name):
        print(f"文件 {file_name} 不存在，无需备份")
        return

    backup_name = f"{os.path.splitext(file_name)[0]}_backup{os.path.splitext(file_name)[1]}"
    try:
        shutil.copy2(file_name, backup_name)
        print(f"已创建备份: {backup_name}")
    except Exception as e:
        print(f"备份文件失败: {e}")

def get_stock_code_with_suffix(stock_code):
    """
    股票代码添加交易所后缀（带后缀即通用格式）
    """
    code = str(stock_code)
    if code.isdigit():
        if len(code) == 6:
            return f"{code}.SH" if code.startswith('6') else f"{code}.SZ"
        if len(code) == 5:
            return f"{code}.HK"
    return code  # 其他情况原样返回

def get_a_share_data():
    try:
        return ak.stock_zh_a_spot_em()
    except Exception as e:
        print(f"拉取A股数据出错: {e}")
        return pd.DataFrame()

def get_hk_share_data():
    try:
        return ak.stock_hk_spot_em()
    except Exception as e:
        print(f"拉取港股数据出错: {e}")
        return pd.DataFrame()

def get_stock_history(stock_code, days=30, max_retries=3):
    import tushare as ts
    pro = ts.pro_api('f1f7dae003add1fcbaa474d3e1cc4d2653bb34c4ca7fb6a597ed87a1')
    end_date = datetime.now().strftime("%Y%m%d")
    start_date = (datetime.now() - timedelta(days=days)).strftime("%Y%m%d")
    code_with_suffix = get_stock_code_with_suffix(stock_code)

    for attempt in range(max_retries):
        try:
            if code_with_suffix.endswith('.HK'):
                result = pro.hk_daily(ts_code=code_with_suffix, start_date=start_date, end_date=end_date)
                time.sleep(30)
            elif code_with_suffix.endswith('.SZ') or code_with_suffix.endswith('.SH'):
                result = pro.daily(ts_code=code_with_suffix, start_date=start_date, end_date=end_date)
                time.sleep(0.5)
            else:
                return pd.DataFrame()

            if result.empty:
                if attempt < max_retries - 1:
                    time.sleep(2)
                    continue
                return pd.DataFrame()

            mapping = {
                'trade_date': '日期',
                'close': '收盘',
                'open': '开盘',
                'high': '最高',
                'low': '最低',
                'change': '涨跌额',
                'pct_chg': '涨跌幅',
                'pre_close': '前收盘'
            }
            cols = {k: v for k, v in mapping.items() if k in result.columns}
            result = result.rename(columns=cols)
            if '前收盘' not in result.columns and '收盘' in result.columns and '涨跌额' in result.columns:
                result['前收盘'] = result['收盘'] - result['涨跌额']
            if '涨跌幅' in result.columns:
                result['涨跌幅'] = result['涨跌幅'] / 100
            if '日期' in result.columns:
                result['日期'] = pd.to_datetime(result['日期'], format='%Y%m%d', errors='coerce')
            result = result.sort_values('日期').reset_index(drop=True)
            return result
        except Exception as e:
            wait = (30 if code_with_suffix.endswith('.HK') else 2 * (attempt + 1))
            print(f"拉取历史数据失败 {stock_code} (第{attempt+1}次): {e}，等待{wait}秒重试")
            time.sleep(wait)
    return pd.DataFrame()

def calculate_volatility(stock_data):
    if not set(['收盘', '涨跌额', '最高', '最低']).issubset(stock_data.columns):
        return None, None, None
    stock_data["前收盘"] = stock_data["收盘"] - stock_data["涨跌额"]
    stock_data["波动率h"] = (stock_data["最高"] - stock_data["前收盘"]) / stock_data["前收盘"]
    stock_data["波动率l"] = (stock_data["最低"] - stock_data["前收盘"]) / stock_data["前收盘"]
    stock_data["负波动率l"] = -stock_data["波动率l"]
    stock_data["波动率"] = stock_data[["波动率h", "负波动率l"]].max(axis=1)
    return (
        stock_data["波动率h"].mean(),
        stock_data["波动率l"].mean(),
        stock_data["波动率"].mean()
    )

def update_excel(data, file_name="hk.xlsx"):
    try:
        wb = openpyxl.load_workbook(file_name)
        ws = wb.active
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(list(data.columns))
    for row in data.itertuples(index=False, name=None):
        ws.append(row)
    wb.save(file_name)
    print(f"已更新 {file_name}")

def update_stock_prices(file_name: str, sheet_name: str):
    """
    用 tushare 拉取行情并更新到excel。不再拉取和写入总股本，仅从sheet读取总股本用于打印。
    """
    import tushare as ts
    pro = ts.pro_api('f1f7dae003add1fcbaa474d3e1cc4d2653bb34c4ca7fb6a597ed87a1')

    print("-----------------------------")
    wb = openpyxl.load_workbook(file_name)
    if sheet_name not in wb.sheetnames:
        print(f"工作表 {sheet_name} 不存在!")
        return False
    ws = wb[sheet_name]
    headers = {cell.value: idx + 1 for idx, cell in enumerate(ws[1])}

    req = ["代码", "现价(CNY)", "现价(HKD)", "今日涨幅", "总股本", "更新时间"]
    for col in req:
        if col not in headers:
            print(f"--- 缺少必须列: {col}")
            return False

    stock_code_col = headers["代码"]
    a_share_price_col = headers["现价(CNY)"]
    hk_share_price_col = headers["现价(HKD)"]
    pct_col = headers["今日涨幅"]
    total_mv_col = headers["总股本"]
    update_time_col = headers["更新时间"]
    previous_low_col = headers.get("前低")

    stock_codes = [
        row[stock_code_col - 1].value
        for row in ws.iter_rows(min_row=2, max_col=stock_code_col + 1)
        if row[stock_code_col].value
    ]
    if not stock_codes:
        print("没有股票代码，无需更新")
        return False

    print("-----------------------------")
    last_hk_call_time = 0
    for i, code in enumerate(stock_codes, start=2):
        ts_code = get_stock_code_with_suffix(code)
        if ts_code is None:
            print(f"--- 无法识别股票代码格式: {code}")
            continue
        now_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            total_stock_issue_val = ws.cell(row=i, column=total_mv_col).value if total_mv_col else None
            if ts_code.endswith('.HK'):
                use_time = time.time()
                if use_time - last_hk_call_time < 30:
                    wait_sec = 30 - (use_time - last_hk_call_time)
                    print(f"等待{wait_sec:.1f}s以避免港股API频率限制")
                    time.sleep(wait_sec)
                df = pro.hk_daily(ts_code=ts_code, limit=1)
                last_hk_call_time = time.time()
                if df.empty:
                    print(f"--- {code} 无法拉取港股行情")
                    continue
                latest = df.sort_values("trade_date", ascending=False).iloc[0]
                price = latest['close']
                pre_close = latest['pre_close']
                pct = (price - pre_close) / pre_close if pre_close else 0
                ws.cell(row=i, column=hk_share_price_col, value=price)
                ws.cell(row=i, column=pct_col, value=pct)
                ws.cell(row=i, column=update_time_col, value=now_time)
                if previous_low_col:
                    curr_low = ws.cell(row=i, column=previous_low_col).value
                    new_low = price if curr_low is None or pd.isna(curr_low) else min(price, curr_low)
                    ws.cell(row=i, column=previous_low_col, value=new_low)
                    print(f"{code:<8} H  {price:>6.2f} {pct*100:>6.2f}% pre_low:{new_low:>6.2f} 股本:{total_stock_issue_val if total_stock_issue_val is not None else 'N/A'}")
                else:
                    print(f"{code:<8} H  {price:>6.2f} {pct*100:>6.2f}% 股本:{total_stock_issue_val if total_stock_issue_val is not None else 'N/A'}")
            elif ts_code.endswith('.SH') or ts_code.endswith('.SZ'):
                df = pro.daily(ts_code=ts_code, limit=1)
                if df.empty:
                    print(f"--- {code} 无法拉取A股行情")
                    continue
                latest = df.sort_values("trade_date", ascending=False).iloc[0]
                price = latest['close']
                pre_close = latest['pre_close']
                pct = (price - pre_close) / pre_close if pre_close else 0
                ws.cell(row=i, column=a_share_price_col, value=price)
                ws.cell(row=i, column=pct_col, value=pct)
                ws.cell(row=i, column=update_time_col, value=now_time)
                if previous_low_col:
                    curr_low = ws.cell(row=i, column=previous_low_col).value
                    new_low = price if curr_low is None or pd.isna(curr_low) else min(price, curr_low)
                    ws.cell(row=i, column=previous_low_col, value=new_low)
                    print(f"{code:<8} A  {price:>6.2f} {pct*100:>6.2f}% pre_low:{new_low:>6.2f} 股本:{total_stock_issue_val if total_stock_issue_val is not None else 'N/A'}")
                else:
                    print(f"{code:<8} A  {price:>6.2f} {pct*100:>6.2f}% 股本:{total_stock_issue_val if total_stock_issue_val is not None else 'N/A'}")
            else:
                print(f"--- 代码无法识别: {code}")
        except Exception as e:
            print(f"行情失败 {code}: {e}")

    ws.cell(row=len(stock_codes) + 5, column=1, value=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    wb.save(file_name)
    print("-----------------------------")
    print(f"已完成价格更新：{file_name}")
    return True

def update_stock_volatility(file_name:str, sheet_name:str, update_prices:bool = True, start_index:int = 1):
    """
    更新波动率与价格。不涉及总股本更新，打印时从sheet读取
    """
    print("-----------------------------")
    print("update stock volatility{}...".format(" and latest closing prices" if update_prices else " only"))

    wb = openpyxl.load_workbook(file_name)
    if sheet_name not in wb.sheetnames:
        print(f"表 {sheet_name} 不存在")
        return
    ws = wb[sheet_name]
    headers = {cell.value: idx+1 for idx, cell in enumerate(ws[1])}

    for need in ["代码", "波动率h", "波动率l", "波动率"]:
        if need not in headers:
            print(f"--- 缺少列: {need}")
            return

    stock_code_col = headers["代码"]
    volh_col = headers["波动率h"]
    voll_col = headers["波动率l"]
    vol_col = headers["波动率"]
    price_col = headers.get("现价(CNY)")
    hk_price_col = headers.get("现价(HKD)")
    pct_col = headers.get("今日涨幅")
    update_time_col = headers.get("更新时间")
    previous_low_col = headers.get("前低")
    total_mv_col = headers.get("总股本")

    stock_codes = [row[stock_code_col-1].value for row in ws.iter_rows(min_row=2, max_col=stock_code_col+1) if row[stock_code_col].value]
    total = len(stock_codes)
    if not total:
        print("没有股票代码")
        return
    if start_index < 1:
        print(f"--- start_index {start_index} < 1，重置为1")
        start_index = 1
    if start_index > total:
        print(f"--- start_index {start_index} 超过股票总数 {total}")
        return

    print("-----------------------------")
    for offset, stock_code in enumerate(stock_codes[start_index-1:], start=start_index):
        excel_row = offset + 1
        scode = str(stock_code)
        stock_data = get_stock_history(scode, days=30)
        total_stock_issue_val = ws.cell(row=excel_row, column=total_mv_col).value if total_mv_col else None
        if stock_data.empty:
            print(f"--- {scode} 无法获取历史数据")
            time.sleep(1)
            continue
        time.sleep(0.5)
        mean_h, mean_l, mean_v = calculate_volatility(stock_data)
        ws.cell(row=excel_row, column=volh_col, value=mean_h)
        ws.cell(row=excel_row, column=voll_col, value=mean_l)
        ws.cell(row=excel_row, column=vol_col, value=mean_v)
        if update_prices and mean_h is not None:
            latest_price = stock_data.iloc[-1]["收盘"]
            pct = stock_data.iloc[-1]["涨跌幅"] / 100
            if len(scode) == 5 and hk_price_col:
                ws.cell(row=excel_row, column=hk_price_col, value=latest_price)
                if pct_col:
                    ws.cell(row=excel_row, column=pct_col, value=pct)
                if previous_low_col:
                    curr_low = ws.cell(row=excel_row, column=previous_low_col).value
                    new_low = latest_price if curr_low is None or pd.isna(curr_low) else min(latest_price, curr_low)
                    ws.cell(row=excel_row, column=previous_low_col, value=new_low)
                    print(f"{scode:<8} H  volatility_h:{mean_h:.4f} volatility_l:{mean_l:.4f} volatility:{mean_v:.4f} price:{latest_price:.2f} change:{pct*100:>6.2f}% pre_low:{new_low:.2f} 股本:{total_stock_issue_val if total_stock_issue_val is not None else 'N/A'}")
                else:
                    print(f"{scode:<8} H  volatility_h:{mean_h:.4f} volatility_l:{mean_l:.4f} volatility:{mean_v:.4f} price:{latest_price:.2f} change:{pct*100:>6.2f}% 股本:{total_stock_issue_val if total_stock_issue_val is not None else 'N/A'}")
            elif len(scode) == 6 and price_col:
                ws.cell(row=excel_row, column=price_col, value=latest_price)
                if pct_col:
                    ws.cell(row=excel_row, column=pct_col, value=pct)
                if previous_low_col:
                    curr_low = ws.cell(row=excel_row, column=previous_low_col).value
                    new_low = latest_price if curr_low is None or pd.isna(curr_low) else min(latest_price, curr_low)
                    ws.cell(row=excel_row, column=previous_low_col, value=new_low)
                    print(f"{scode:<8} A  volatility_h:{mean_h:.4f} volatility_l:{mean_l:.4f} volatility:{mean_v:.4f} price:{latest_price:.2f} change:{pct*100:>6.2f}% pre_low:{new_low:.2f} 股本:{total_stock_issue_val if total_stock_issue_val is not None else 'N/A'}")
                else:
                    print(f"{scode:<8} A  volatility_h:{mean_h:.4f} volatility_l:{mean_l:.4f} volatility:{mean_v:.4f} price:{latest_price:.2f} change:{pct*100:>6.2f}% 股本:{total_stock_issue_val if total_stock_issue_val is not None else 'N/A'}")
            else:
                print(f"{scode:<8}    volatility_h:{mean_h:.4f} volatility_l:{mean_l:.4f} volatility:{mean_v:.4f} 股本:{total_stock_issue_val if total_stock_issue_val is not None else 'N/A'}")
            if update_time_col:
                ws.cell(row=excel_row, column=update_time_col, value=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        else:
            print(f"{scode:<8}    volatility_h:{mean_h:.4f} volatility_l:{mean_l:.4f} volatility:{mean_v:.4f} 股本:{total_stock_issue_val if total_stock_issue_val is not None else 'N/A'}")

    wb.save(file_name)
    print("-----------------------------")
    print(f"{'已完成波动率与价格更新' if update_prices else '已完成波动率更新'}：{file_name}")

def main():
    parser = argparse.ArgumentParser(description="update stock data")
    parser.add_argument('-p', '--price', action='store_true', help="仅更新股价")
    parser.add_argument('-v', '--volatility', action='store_true', help="仅更新波动率")
    parser.add_argument('-a', '--all', action='store_true', help="全部更新")
    parser.add_argument('-s', '--start', type=int, default=1, help="起始行号(用于波动率更新)")

    args = parser.parse_args()

    file_name = "ValueInvestment_auto.xlsx"
    sheet_name = "预期收益率管理"

    create_backup(file_name)
    if args.all:
        price_updated = update_stock_prices(file_name, sheet_name)
        update_stock_volatility(file_name, sheet_name, update_prices=not price_updated, start_index=args.start)
    elif args.price:
        update_stock_prices(file_name, sheet_name)
    elif args.volatility:
        update_stock_volatility(file_name, sheet_name, update_prices=True, start_index=args.start)
    else:
        update_stock_prices(file_name, sheet_name)

    try:
        os.system(f"open \"{file_name}\"")
    except Exception:
        pass

if __name__ == "__main__":
    main()

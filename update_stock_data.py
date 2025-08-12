import akshare as ak
import openpyxl
import pandas as pd
from datetime import datetime, timedelta
import os
import argparse

def get_a_share_data():
    """
    get the stock data of A-shares
    """
    return ak.stock_zh_a_spot_em()

def get_hk_share_data():
    """
    get the stock data of HK-shares
    """
    return ak.stock_hk_spot_em()

def get_stock_history(stock_code, days=30):
    """
    get the stock history data
    @stock_code: str, stock code
    @days: int, the number of days
    """
    end_date = datetime.now().strftime("%Y%m%d")
    start_date = (datetime.now() - timedelta(days=days)).strftime("%Y%m%d")
    if len(stock_code) == 5:
        return ak.stock_hk_hist(symbol=stock_code,period = "daily",start_date=start_date, end_date=end_date, adjust="qfq")
    elif len(stock_code) == 6:
        return ak.stock_zh_a_hist(symbol=stock_code, period="daily",start_date=start_date, end_date=end_date, adjust="qfq")
    else:
        return pd.DataFrame()
    
def calculate_volatility(stock_data):
    """
    calculate the volatility of stock
    @stock_data: pandas.DataFrame, stock data
    """
    stock_data["前收盘"] = stock_data["收盘"] - stock_data["涨跌额"]
    stock_data["波动率h"] = (stock_data["最高"] - stock_data["前收盘"]) / stock_data["前收盘"]
    stock_data["波动率l"] = (stock_data["最低"] - stock_data["前收盘"]) / stock_data["前收盘"]
    stock_data["负波动率l"] = -stock_data["波动率l"]
    stock_data["波动率"] = stock_data[["波动率h", "负波动率l"]].max(axis=1)
    mean_volatility_h = stock_data["波动率h"].mean()
    mean_volatility_l = stock_data["波动率l"].mean()
    mean_volatility = stock_data["波动率"].mean()
    return mean_volatility_h, mean_volatility_l, mean_volatility

def update_excel(data, file_name = "hk.xlsx"):
    """
    save the stock data to excel
    @data: pandas.DataFrame, stock data
    @file_name: excel name
    """
    try:
        wb = openpyxl.load_workbook(file_name)
        ws = wb.active
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        ws = wb.active
        # write header
        headers = list(data.columns)
        ws.append(headers)

    # write data
    for row in data.itertuples(index=False, name=None):
        ws.append(row)

    wb.save(file_name)
    print(f"data is updated in {file_name}")

def update_stock_prices(file_name:str, sheet_name:str):
    """
    update stock prices in exist excel
    """
    print("-----------------------------")
    # read excel
    wb = openpyxl.load_workbook(file_name)
    if sheet_name not in wb.sheetnames:
        print(f"sheet {sheet_name} is not exist!")
        return
    
    ws = wb[sheet_name]

    headers = {cell.value: idx+1 for idx, cell in enumerate(ws[1])}  # 标题 -> 列号映射

    required_columns = ["代码", "现价(CNY)", "现价(HKD)", "今日涨幅", "总股本", "更新时间"]
    for col in required_columns:
        if col not in headers:
            print(f"--- Error!!! --- column {col} is missing.")
            return

    stock_code_col = headers["代码"]
    a_share_price_col = headers["现价(CNY)"]
    hk_share_price_col = headers["现价(HKD)"]
    percentage_change_col = headers["今日涨幅"]
    total_stock_issue_col = headers["总股本"]
    update_time_col = headers["更新时间"]

    stock_codes = [row[stock_code_col-1].value for row in ws.iter_rows(min_row=2, max_col=stock_code_col+1) if row[stock_code_col].value]

    # fetching share data
    print("fetching A-share data...")
    a_share_data = get_a_share_data()

    print("fetching HK-share data...")
    hk_share_data = get_hk_share_data()

    # for debug
    debug_flag = False
    if debug_flag:
        update_excel(a_share_data, file_name = "a.xlsx")
        update_excel(a_share_data, file_name = "hk.xlsx")

    print("-----------------------------")
    # update stock prices
    for i, stock_code in enumerate(stock_codes, start=2):
        stock_code = str(stock_code)
        if len(stock_code) == 5:  # hk stock
            row_data = hk_share_data[hk_share_data["代码"] == stock_code]
            if not row_data.empty:
                company_name = row_data.iloc[0]["名称"]
                latest_price = row_data.iloc[0]["最新价"]
                percentage_change = row_data.iloc[0]["涨跌幅"] / 100
                ws.cell(row=i, column=hk_share_price_col, value=latest_price) # write hk stock price
                ws.cell(row=i, column=percentage_change_col, value=percentage_change) # write hk stock percentage change
                ws.cell(row=i, column=update_time_col, value=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                print(f"{stock_code:<8} {'H':<2} {company_name:<12} {latest_price:>6.2f} {percentage_change*100:>6.2f}%")
            else:
                print(f"--- Warning!!! --- {stock_code} is not found.")
        elif len(stock_code) == 6:   # A stock
            row_data = a_share_data[a_share_data["代码"] == stock_code]
            if not row_data.empty:
                company_name = row_data.iloc[0]["名称"]
                latest_price = row_data.iloc[0]["最新价"]
                percentage_change = row_data.iloc[0]["涨跌幅"] / 100
                total_val = row_data.iloc[0]["总市值"]
                total_stock_issue = total_val / latest_price * 1e-8
                if stock_code == "600025":
                    total_stock_issue = 188.3 # special case for 600025
                ws.cell(row=i, column=a_share_price_col, value=latest_price)
                ws.cell(row=i, column=total_stock_issue_col, value=total_stock_issue)
                ws.cell(row=i, column=percentage_change_col, value=percentage_change) # write hk stock percentage change
                ws.cell(row=i, column=update_time_col, value=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                print(f"{stock_code:<8} {'A':<2} {company_name:<12} {latest_price:>6.2f} {percentage_change*100:>6.2f}% total_stock_issue:{total_stock_issue:.2f}")
            else:
                print(f"--- Warning!!! --- {stock_code} is not found.")

    ws.cell(row=len(stock_codes)+5, column=1, value=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    wb.save(file_name)
    print("-----------------------------")
    print(f"the stock prices are updated in {file_name}.")

def update_stock_volatility(file_name:str, sheet_name:str):
    """
    update stock volatility in exist excel
    """
    print("-----------------------------")
    print("update stock volatility...")

    # read excel
    wb = openpyxl.load_workbook(file_name)
    if sheet_name not in wb.sheetnames:
        print(f"sheet {sheet_name} is not exist!")
        return
    
    ws = wb[sheet_name]

    headers = {cell.value: idx+1 for idx, cell in enumerate(ws[1])}  # 标题 -> 列号映射

    required_columns = ["代码", "波动率h", "波动率l", "波动率"]
    for col in required_columns:
        if col not in headers:
            print(f"--- Error!!! --- column {col} is missing.")
            return

    stock_code_col = headers["代码"]
    volatility_h_col = headers["波动率h"]
    volatility_l_col = headers["波动率l"]
    volatility_col = headers["波动率"]

    stock_codes = [row[stock_code_col-1].value for row in ws.iter_rows(min_row=2, max_col=stock_code_col+1) if row[stock_code_col].value]

    print("-----------------------------")
    # update stock volatility
    for i, stock_code in enumerate(stock_codes, start=2):
        stock_code = str(stock_code)
        stock_data = get_stock_history(stock_code, days=30)
        if stock_data.empty:
            print(f"--- Warning!!! --- {stock_code} is not found.")
            continue
        mean_volatility_h, mean_volatility_l, mean_volatility = calculate_volatility(stock_data)
        ws.cell(row=i, column=volatility_h_col, value=mean_volatility_h)
        ws.cell(row=i, column=volatility_l_col, value=mean_volatility_l)
        ws.cell(row=i, column=volatility_col, value=mean_volatility)
        print(f"{stock_code:<8}  volatility_h:{mean_volatility_h:.4f}  volatility_l:{mean_volatility_l:.4f}  volatility:{mean_volatility:.4f}")

    wb.save(file_name)
    print("-----------------------------")
    print(f"the stock volatility are updated in {file_name}.")

def main():
    parser = argparse.ArgumentParser(description="Update stock data")
    parser.add_argument('-p', '--price', action='store_true', help="Only update stock prices.")
    parser.add_argument('-v', '--volatility', action='store_true', help="Only update stock volatility.")
    parser.add_argument('-a', '--all', action='store_true', help="Update both stock prices and volatility.")

    args = parser.parse_args()

    file_name = "ValueInvestment_auto.xlsx"
    sheet_name = "预期收益率管理"

    if args.all:
        update_stock_prices(file_name, sheet_name)
        update_stock_volatility(file_name, sheet_name)
    elif args.price:
        update_stock_prices(file_name, sheet_name)
    elif args.volatility:
        update_stock_volatility(file_name, sheet_name)
    else:
        update_stock_prices(file_name, sheet_name)

    os.system(f"open {file_name}")

if __name__ == "__main__":
    main()
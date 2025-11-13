# Investment Tools 说明

## 工具用法简介

### 1. update_stock_data.py
使用 akshare 库更新股票价格和波动率数据到 Excel 文件中。支持A股和港股，可选择只更新价格，只更新波动率或全部更新。
推荐
```bash
python update_stock_data.py -a
python update_stock_data.py -v -s $start_idx
python update_stock_data.py -p --homemade
```

### 2. update_stock_data_tushare.py
类似 update_stock_data.py，但使用 tushare 作为数据源。

### 3. stock_quote_scraper.py
抓取股票实时报价，可通过命令行指定股票代码，支持批量查询并自动控制请求频率。
```bash
python stock_quote_scraper.py $code
```

### 4. backtest/backtest_analysis_and_valuation.ipynb
根据ETF历史价格趋势，进行技术分析和估值计算。
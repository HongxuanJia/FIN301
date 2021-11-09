import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

def read_excel(filepath):
    statement = pd.read_excel(filepath)
    year = statement["Data Year - Fiscal"]
    current_asset = statement["Current Assets - Total"]
    current_liability = statement["Current Liabilities - Total"]
    long_term_debt = statement["Long-Term Debt - Total"]
    total_asset = statement["Assets - Total"]
    total_liability = statement["Liabilities - Total"]
    total_shareholder_equity = statement["Stockholders Equity - Total"]
    net_income = statement["Net Income (Loss)"]
    preferred_Dividend = statement["Dividends - Preferred/Preference"]
    number_outstanding = statement["Common Shares Outstanding"]
    stock_price = statement["Price Close - Annual - Calendar"]
    cost_goods = statement["Cost of Goods Sold"]
    inventory = statement["Inventories - Total"]
    # credit_sales = statement[""]
    account_receivable = statement["Receivables - Total"]
    sales = statement["Sales/Turnover (Net)"]
    working_capital = statement["Working Capital (Balance Sheet)"]
    # supplier_purchase = statement[""]
    account_payable = statement["Accounts Payable - Trade"]
    EPS_basic = statement["Earnings Per Share (Basic) - Including Extraordinary Items"]
    EPS_diluted = statement["Earnings Per Share (Diluted) - Including Extraordinary Items"]
    dic = {
        "year" : year,
        "current_asset" : current_asset,
        "current_liability" : current_liability,
        "long_term_debt" : long_term_debt,
        "total_asset" : total_asset,
        "total_liability" : total_liability,
        "totol_shareholder_equity" : total_shareholder_equity,
        "net_income" : net_income,
        "preferred_Dividend" : preferred_Dividend,
        "number_outstanding" : number_outstanding,
        "stock_price" : stock_price,
        "cost_goods" : cost_goods,
        "inventory" : inventory,
        # "credit_sales" : credit_sales,
        "account_receivable" : account_receivable,
        "sales" : sales,
        "working_capital" : working_capital,
        # "supplier_purchase" : supplier_purchase,
        "account_payable" : account_payable,
        "EPS_basic" : EPS_basic,
        "EPS_diluted" : EPS_diluted
    }
    variables = pd.DataFrame(dic)
    return variables

def ratios(variables):
    # variables[""]
    # 1. 公司短期变现能力
    Liquidity_ratio = variables["current_asset"]/variables["current_liability"]
    Quick_ratio = (variables["current_asset"]-variables["inventory"])/variables["current_liability"]
    # 2. 长期偿债能力分析
    Leverage = variables["total_liability"]/variables["current_asset"]
    Debt_Equity_ratio = variables["total_liability"]/variables["totol_shareholder_equity"]
    # 3. 投资收益分析
    EPS = variables["EPS_basic"]
    P_E_ratio = variables["stock_price"]/variables["EPS_basic"]
    # 4. 流动资产周转率分析
    Inventory_turnover = variables["cost_goods"]/variables["inventory"]
    Capital_employed_turnover = variables["sales"]/(variables["totol_shareholder_equity"]+variables["long_term_debt"])
    Working_capital = variables["sales"]/variables["working_capital"]
    Asset_turnover = variables["sales"]/variables["total_asset"]
    dic = {
        "year": variables["year"],
        "Liquidity_ratio":Liquidity_ratio,
        "Quick_ratio":Quick_ratio,
        "Leverage":Leverage,
        "Debt_Equity_ratio":Debt_Equity_ratio,
        "EPS":EPS,
        "P_E_ratio":P_E_ratio,
        "Inventory_turnover":Inventory_turnover,
        "Capital_employed_turnover":Capital_employed_turnover,
        "Working_capital_turnover":Working_capital,
        "Asset_turnover":Asset_turnover
    }
    ratios = pd.DataFrame(dic)
    return ratios

if __name__ == '__main__':
    Pfizer_ratios = ratios(read_excel("./Statements/Pfizer_2000-2020.xlsx"))
    Pfizer_ratios.to_excel("Ratios/Pfizer_Ratios.xlsx", encoding="utf-8")
    Medivation_ratios = ratios(read_excel("Statements/Medivation_2005-2015.xlsx"))
    Medivation_ratios.to_excel("Ratios/Medivation_Ratios.xlsx", encoding="utf-8")
    Hospira_ratios = ratios(read_excel("Statements/Hospira_2002-2014.xlsx"))
    Hospira_ratios.to_excel("Ratios/Hospira_Ratios.xlsx", encoding="utf-8")
    Wyeth_ratios = ratios(read_excel("Statements/Wyeth_2001-2008.xlsx"))
    Wyeth_ratios.to_excel("Ratios/Wyeth_Ratios.xlsx", encoding="utf-8")
    # read_excel("Statements/Medivation_2005-2015.xlsx")
    # read_excel("Statements/Hospira_2002-2014.xlsx")
    # read_excel("Statements/Wyeth_2001-2008.xlsx")

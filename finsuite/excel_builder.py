# -*- coding: utf-8 -*-
from dataclasses import dataclass
from datetime import date, timedelta
import pandas as pd

@dataclass
class Options:
    fy_start: str
    fy_end: str
    currency: str
    rtl: bool = True

def _today(): return date.today().isoformat()

def sheet_README(opts): 
    return pd.DataFrame([["FinSuite Pro — اکسل‌ساز حسابداری (PyQt6)"],
                         ["سال مالی", f"{opts.fy_start} تا {opts.fy_end}"],
                         ["واحد پول", opts.currency],
                         [""],["راهنما"],
                         ["COA/MASTER_DATA را کامل و JOURNAL را وارد کنید."]],
                        columns=["اطلاعات"])

def sheet_CONFIG(opts):
    return pd.DataFrame([["FY_START",opts.fy_start,""],["FY_END",opts.fy_end,""],["CURRENCY_GROUP",opts.currency,""],["RTL_TITLES","True",""]],
                        columns=["Key","Value","Notes"])

def sheet_COA():
    return pd.DataFrame([["100101","صندوق","3","Assets","Debit","","1001",True,""],
                         ["111101","بانک جاری","4","Assets","Debit","","100101",True,""],
                         ["411101","درآمد فروش","3","Revenue","Credit","","4000",True,""],
                         ["511101","بهای تمام شده","3","COGS","Debit","","5000",True,""]],
                        columns=["AccountCode","AccountName","Level","Group","NormalSide","Currency","Parent","IsActive","Notes"])

def sheet_MASTER_DATA():
    return pd.DataFrame([["PARTY","CUST-A","مشتری الف","",""],["PARTY","VEND-B","تامین‌کننده ب","",""]],
                        columns=["Type","Code","Name","Parent","Notes"])

def sheet_JOURNAL(opts):
    return pd.DataFrame([[opts.fy_start,"JV-0001","INV-1001","ثبت فروش","411101","درآمد فروش","","","CUST-A",opts.currency,1,"",15000000,9,"=L2*0.09","","","C1","CUST-A",""],
                         [opts.fy_start,"JV-0001","INV-1001","بانک","111101","بانک","","","CUST-A",opts.currency,1,16350000,"","","","","C1","CUST-A",""]],
                        columns=["Date","JrnNo","DocNo","Description","AccountCode","AccountName","CostCenter","Project","Partner","Currency","FXRate","Debit","Credit","VATRate","VATAmount","WHTRate","WHTAmount","Entity","Counterparty","Ref"])

def sheet_GL():
    return pd.DataFrame([["111101","بانک",_today(),"JV-0001","INV-1001","وصول فروش","",16350000,"=SUM(G2:H2*-1)"],
                         ["411101","درآمد فروش",_today(),"JV-0001","INV-1001","ثبت فروش",15000000,"","=SUM(G3:H3*-1)"]],
                        columns=["AccountCode","AccountName","Date","JrnNo","DocNo","Description","Debit","Credit","Balance"])

def sheet_TRIAL():
    return pd.DataFrame([["111101","بانک",0,0,0,16350000,"=MAX(0,(C2+E2)-(D2+F2))","=MAX(0,(D2+F2)-(C2+E2))"],
                         ["411101","درآمد فروش",0,0,15000000,0,"=MAX(0,(C3+E3)-(D3+F3))","=MAX(0,(D3+F3)-(C3+E3))"]],
                        columns=["AccountCode","AccountName","OpeningDr","OpeningCr","PeriodDr","PeriodCr","ClosingDr","ClosingCr"])

def sheet_PL():
    return pd.DataFrame([["REV","فروش خالص","=SUMIFS(TRIAL!E:E,TRIAL!A:A,\">=400000\")-SUMIFS(TRIAL!F:F,TRIAL!A:A,\">=400000\")","=C2"],
                         ["COGS","بهای تمام‌شده","=0","=C3"],
                         ["GP","سود ناخالص","=C2-C3","=C4"]],
                        columns=["LineCode","LineName","Formula","Amount"])

def sheet_BS():
    return pd.DataFrame([["CASH","وجه نقد","=SUMIFS(TRIAL!G:G,TRIAL!A:A,111101)-SUMIFS(TRIAL!H:H,TRIAL!A:A,111101)","=C2"],
                         ["EQUITY","حقوق صاحبان سهام","0","0"]],
                        columns=["LineCode","LineName","Formula","Amount"])

def sheet_CASHFLOW():
    return pd.DataFrame([["CFO","CF-01","جریان نقد عملیاتی","0"]],
                        columns=["Section","Line","Description","Amount"])

def sheet_SALES_PURCH():
    return pd.DataFrame([["SALE", _today(),"INV-1001","مشتری الف","فروش نمونه",15000000,9,"=F2*G2/100","=F2+H2","411101"]],
                        columns=["Type","Date","DocNo","Counterparty","Description","TaxBase","VATRate","VATAmount","Total","AccountCode"])

def sheet_VAT_RETURN():
    return pd.DataFrame([["1404-01","=SUMIFS(SALES_PURCH!H:H,SALES_PURCH!A:A,\"SALE\")","=0","=B2-C2"]],
                        columns=["Period","OutputVAT","InputVAT","Payable"])

def sheet_AP_AR():
    due = (date.today()+timedelta(days=30)).isoformat()
    return pd.DataFrame([["CUSTOMER","CUST-A","مشتری الف","INV-1001",_today(),due,16350000,0,"=G2-H2"]],
                        columns=["Type","Code","Name","InvoiceNo","InvoiceDate","DueDate","Amount","Received(Paid)","Balance"])

def sheet_BANK_REC():
    return pd.DataFrame([[ _today(),"STM-001","نمونه",1000000,1000000,"=IF(D2=E2,\"✔\",\"\")","=D2-E2"]],
                        columns=["Date","StatementRef","Description","Bank(+/-)","GL(+/-)","Tick","Difference"])

def sheet_INVENTORY():
    return pd.DataFrame([["SKU-1","کالای نمونه",_today(),"RCV-1",10,0,100,10,100,1000]],
                        columns=["SKU","ItemName","Date","DocNo","InQty","OutQty","UnitCost","RunQty","RunAvgCost","RunValue"])

def sheet_FIXED_ASSETS():
    return pd.DataFrame([["FA-001","رایانه",_today(),60000000,"SL",36,0,"=IF(E2=\"SL\",(D2-G2)/F2,\"\")","=SUM($H$2:H2)","=D2-I2"]],
                        columns=["AssetID","AssetName","PurchaseDate","Cost","Method","Life(Months)","Salvage","MonthlyDep","AccumDep","NBV"])

def sheet_PAYROLL():
    return pd.DataFrame([["E-001","کارمند نمونه",120000000,10000000,0,"=SUM(C2:E2)","=ROUND(0.07*F2,0)","=ROUND(0.10*F2,0)",0,"=F2-G2-H2-I2","=ROUND(0.23*F2,0)"]],
                        columns=["EmpID","Name","BaseSalary","Allowances","Overtime","Gross","Pension_EE","Tax","OtherDeductions","Net","Pension_ER"])

def sheet_BUDGET():
    return pd.DataFrame([["1404-01","411101","درآمد فروش",18000000,15000000,"=D2-E2"]],
                        columns=["Period","AccountCode","AccountName","Budget","Actual","Variance"])

def sheet_FX_RATES(opts):
    return pd.DataFrame([[opts.fy_start,opts.currency,1]], columns=["Date","Currency","RateToGroup"])

def sheet_ENTITIES():
    return pd.DataFrame([["C1","شرکت مادر","IRR",100,"IR"]],
                        columns=["EntityCode","EntityName","Currency","Ownership%","Country"])

def sheet_MAP_PL():
    return pd.DataFrame([["411101","REV",""]], columns=["AccountCode","PL_LineCode","Notes"])

def sheet_MAP_BS():
    return pd.DataFrame([["111101","CASH",""]], columns=["AccountCode","BS_LineCode","Notes"])

def sheet_CONS_ELIMS():
    return pd.DataFrame([[ _today(),"ELIM-1","معاملات درون‌گروهی","S1","411101","S1","511101",5000000,"IRR" ]],
                        columns=["Date","ElimNo","Description","EntityDr","AccountDr","EntityCr","AccountCr","Amount","Currency"])

def sheet_TB_GROUP_PRE():
    return pd.DataFrame([["111101","بانک",0,0,0,16350000,"=MAX(0,(C2+E2)-(D2+F2))","=MAX(0,(D2+F2)-(C2+E2))"]],
                        columns=["AccountCode","AccountName","OpeningDr","OpeningCr","PeriodDr","PeriodCr","ClosingDr","ClosingCr"])

def sheet_TB_GROUP_POST():
    return pd.DataFrame([["111101","بانک","=E2","=F2","نمونه"]],
                        columns=["AccountCode","AccountName","Adj_PeriodDr","Adj_PeriodCr","Note"])

def sheet_CLOSE_CHECK():
    return pd.DataFrame([["1","ثبت اسناد","",""]], columns=["Step","شرح","مسئول","وضعیت"])

def sheet_QUALITY_KPIs():
    return pd.DataFrame([["Journal Balanced","=IF(SUM(JOURNAL!L:L)=SUM(JOURNAL!M:M),\"OK\",\"ERROR\")","=B2"]],
                        columns=["Check","Formula","Result"])

MODULES = {
    "README": sheet_README, "CONFIG": sheet_CONFIG, "COA": sheet_COA, "MASTER_DATA": sheet_MASTER_DATA,
    "JOURNAL": sheet_JOURNAL, "GL": sheet_GL, "TRIAL": sheet_TRIAL, "PL": sheet_PL, "BS": sheet_BS,
    "CASHFLOW": sheet_CASHFLOW, "SALES_PURCH": sheet_SALES_PURCH, "VAT_RETURN": sheet_VAT_RETURN,
    "AP_AR": sheet_AP_AR, "BANK_REC": sheet_BANK_REC, "INVENTORY": sheet_INVENTORY,
    "FIXED_ASSETS": sheet_FIXED_ASSETS, "PAYROLL": sheet_PAYROLL, "BUDGET": sheet_BUDGET,
    "FX_RATES": sheet_FX_RATES, "ENTITIES": sheet_ENTITIES, "MAP_PL": sheet_MAP_PL, "MAP_BS": sheet_MAP_BS,
    "CONS_ELIMS": sheet_CONS_ELIMS, "TB_GROUP_PRE": sheet_TB_GROUP_PRE, "TB_GROUP_POST": sheet_TB_GROUP_POST,
    "CLOSE_CHECK": sheet_CLOSE_CHECK, "QUALITY_KPIs": sheet_QUALITY_KPIs,
}

def build_excel(path, selected_keys, opts):
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        for key in selected_keys:
            MODULES[key](opts).to_excel(writer, sheet_name=key[:31], index=False)
    return path

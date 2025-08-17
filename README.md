# FinSuite Pro — نرم‌افزار دسکتاپ اکسل‌ساز حسابداری (PyQt6)

## اجرا
1) نصب وابستگی‌ها:
```bash
pip install -r requirements.txt
```
2) اجرا:
```bash
python run.py
```

## ساخت exe (اختیاری)
```bash
pip install pyinstaller
pyinstaller --noconsole --onefile --add-data "finsuite/resources;finsuite/resources" run.py -n FinSuitePro
```

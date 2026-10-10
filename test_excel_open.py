try:
    import win32com.client
    import time

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    t0 = time.time()
    wb = excel.Workbooks.Open(r"C:\Users\MaximuM-Tech\Downloads\بنك_اسئلة_القسم_المتوسط_الضبع_الاسود_LMS.xlsx")
    t1 = time.time()
    print(f"Excel opened Medium Bank in: {t1 - t0:.2f} seconds")
    wb.Close(False)

    t0 = time.time()
    wb2 = excel.Workbooks.Open(r"C:\Users\MaximuM-Tech\Downloads\بنك_اسئلة_القسم_النهائي_الضبع_الاسود_LMS.xlsx")
    t1 = time.time()
    print(f"Excel opened Final Bank in: {t1 - t0:.2f} seconds")
    wb2.Close(False)

    excel.Quit()
except Exception as e:
    print(f"Error testing Excel COM: {e}")

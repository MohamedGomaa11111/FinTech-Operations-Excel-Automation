import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
try:
    xlsm_path = r"c:\Users\mohamed\Desktop\xlsx project\Cash_Transfer_System.xlsm"
    wb = excel.Workbooks.Open(xlsm_path)
    ws_dash = wb.Sheets("لوحة التحكم")
    ws_acc = wb.Sheets("إدارة الحسابات")
    
    # Clean any existing charts first
    for i in range(ws_dash.ChartObjects().Count, 0, -1):
        ws_dash.ChartObjects(i).Delete()
        
    print("Adding Chart 1: Liquidity Distribution Donut...")
    ch1_shape = ws_dash.Shapes.AddChart2(297, 80, 480, 240, 360, 230)
    ch1 = ch1_shape.Chart
    ch1.SetSourceData(ws_dash.Range("I13:J17"))
    ch1.HasTitle = True
    ch1.ChartTitle.Text = "توزيع السيولة النقدية ورأس المال"
    print("Chart 1 created!")
    
    print("Adding Chart 2: Wallet Lines Bar Chart...")
    # Using Union of Name and Balance from tbl_Accounts
    rng_names = ws_acc.Range("C5:C12") # top 8 wallet lines
    rng_bals = ws_acc.Range("J5:J12")
    rng_union = excel.Union(rng_names, rng_bals)
    
    ch2_shape = ws_dash.Shapes.AddChart2(201, 57, 50, 240, 410, 230)
    ch2 = ch2_shape.Chart
    ch2.SetSourceData(rng_union)
    ch2.HasTitle = True
    ch2.ChartTitle.Text = "مقارنة أرصدة خطوط المحافظ الرئيسية"
    ch2.HasLegend = False
    print("Chart 2 created!")
    
    wb.Save()
    # Update XLSX copy too
    xlsx_final = r"c:\Users\mohamed\Desktop\xlsx project\Cash_Transfer_System.xlsx"
    wb.SaveCopyAs(xlsx_final)
    wb.Close(False)
    print("Workbook and charts successfully updated!")
except Exception as e:
    print("Chart creation error:", e)
finally:
    excel.Quit()

# -*- coding: utf-8 -*-
"""
Advanced Excel Builder & VBA Injector.
Uses win32com to:
1. Open Cash_Transfer_System_Base.xlsx
2. Add Excel Data Validations (dropdowns)
3. Inject complete VBA automation macros
4. Insert functional interactive buttons
5. Add native Dashboard charts
6. Save as Cash_Transfer_System.xlsm and Cash_Transfer_System.xlsx
"""

import os
import sys
import win32com.client

VBA_CODE = '''
Option Explicit

' =========================================================================
' Professional Cash Transfer & Financial Management System - Core Automation
' Accounting Engine, Ledger Posting, Reversals, and Daily Closing
' =========================================================================

Public Sub SaveTransaction()
    On Error GoTo ErrHandler
    Dim wsEntry As Worksheet, wsTxns As Worksheet, wsLedger As Worksheet, wsAcc As Worksheet
    Dim tblTxn As ListObject, tblLed As ListObject
    Dim newTxnRow As ListRow, newLedRow As ListRow
    
    Set wsEntry = ThisWorkbook.Sheets("إدخال عملية")
    Set wsTxns = ThisWorkbook.Sheets("سجل العمليات")
    Set wsLedger = ThisWorkbook.Sheets("دفتر الحركات")
    Set wsAcc = ThisWorkbook.Sheets("إدارة الحسابات")
    
    Set tblTxn = wsTxns.ListObjects("tbl_Transactions")
    Set tblLed = wsLedger.ListObjects("tbl_Ledger")
    
    ' Read Form Values
    Dim txnID As String, txnDate As String, txnTime As String, txnType As String
    Dim custName As String, custPhone As String
    Dim srcAcc As String, dstAcc As String, feeAcc As String
    Dim pAmount As Double, rAmount As Double, dAmount As Double
    Dim comm As Double, extFee As Double, dirExp As Double, netProf As Double
    Dim notes As String, status As String
    
    txnID = Trim(wsEntry.Range("D4").Value)
    If txnID = "" Then
        txnID = "TXN-" & Format(Date, "yyyymmdd") & "-" & Format(Int(Rnd() * 9000 + 1000), "0000")
        wsEntry.Range("D4").Value = txnID
    End If
    
    txnDate = Format(wsEntry.Range("D5").Value, "yyyy-mm-dd")
    txnTime = wsEntry.Range("D6").Text
    txnType = Trim(wsEntry.Range("D7").Value)
    custName = Trim(wsEntry.Range("D8").Value)
    custPhone = Trim(wsEntry.Range("D9").Value)
    
    srcAcc = Trim(wsEntry.Range("G4").Value)
    dstAcc = Trim(wsEntry.Range("G5").Value)
    pAmount = Val(wsEntry.Range("G6").Value)
    rAmount = Val(wsEntry.Range("G7").Value)
    dAmount = Val(wsEntry.Range("G8").Value)
    
    comm = Val(wsEntry.Range("J4").Value)
    extFee = Val(wsEntry.Range("J5").Value)
    feeAcc = Trim(wsEntry.Range("J6").Value)
    dirExp = Val(wsEntry.Range("J7").Value)
    netProf = comm - extFee - dirExp
    
    notes = Trim(wsEntry.Range("D11").Value)
    status = Trim(wsEntry.Range("G11").Value)
    If status = "" Then status = "ناجحة"
    
    ' 1. VALIDATIONS
    If txnType = "" Then
        MsgBox "خطأ: يرجى تحديد نوع العملية!", vbCritical, "نظام التحويلات المالية"
        Exit Sub
    End If
    If custName = "" Then
        MsgBox "خطأ: يرجى إدخال اسم العميل!", vbCritical, "نظام التحويلات المالية"
        Exit Sub
    End If
    If pAmount <= 0 And txnType <> "أخرى" Then
        MsgBox "خطأ: المبلغ الأساسي للتحويل يجب أن يكون أكبر من صفر!", vbCritical, "نظام التحويلات المالية"
        Exit Sub
    End If
    If comm < 0 Or extFee < 0 Or dirExp < 0 Then
        MsgBox "خطأ: لا يمكن إدخال قيم سالبة في العمولات أو الرسوم!", vbCritical, "نظام التحويلات المالية"
        Exit Sub
    End If
    If extFee > 0 And feeAcc = "" Then
        feeAcc = srcAcc
        wsEntry.Range("J6").Value = feeAcc
    End If
    
    ' Check Duplicate Transaction ID
    Dim checkCell As Range
    Set checkCell = tblTxn.ListColumns(1).DataBodyRange.Find(What:=txnID, LookIn:=xlValues, LookAt:=xlWhole)
    If Not checkCell Is Nothing Then
        MsgBox "خطأ: رقم العملية (" & txnID & ") مسجل مسبقاً لمنع التكرار!", vbCritical, "تنبيه تكرار"
        Exit Sub
    End If
    
    ' 2. APPEND TO tbl_Transactions
    Application.ScreenUpdating = False
    Set newTxnRow = tblTxn.ListRows.Add
    With newTxnRow
        .Range(1).Value = txnID
        .Range(2).Value = txnDate
        .Range(3).Value = txnTime
        .Range(4).Value = txnType
        .Range(5).Value = custName
        .Range(6).Value = custPhone
        .Range(7).Value = srcAcc
        .Range(8).Value = dstAcc
        .Range(9).Value = pAmount
        .Range(10).Value = rAmount
        .Range(11).Value = dAmount
        .Range(12).Value = comm
        .Range(13).Value = extFee
        .Range(14).Value = feeAcc
        .Range(15).Value = dirExp
        .Range(16).Value = netProf
        .Range(17).Value = status
        .Range(18).Value = notes
        .Range(19).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    End With
    
    ' 3. GENERATE ATOMIC FINANCIAL MOVEMENTS IN tbl_Ledger
    Dim movCount As Long
    movCount = tblLed.ListRows.Count
    
    Select Case txnType
        Case "كاش إلى محفظة"
            ' Movement 1: Cash received from customer into cash drawer
            Set newLedRow = tblLed.ListRows.Add
            movCount = movCount + 1
            newLedRow.Range(1).Value = "MOV-" & Format(movCount, "000000")
            newLedRow.Range(2).Value = txnID
            newLedRow.Range(3).Value = txnDate
            newLedRow.Range(4).Value = txnTime
            newLedRow.Range(5).Value = srcAcc
            newLedRow.Range(6).Value = GetAccountName(srcAcc)
            newLedRow.Range(7).Value = "إيداع نقدي (عميل)"
            newLedRow.Range(8).Value = "استلام نقدية من العميل " & custName
            newLedRow.Range(9).Value = rAmount ' Incoming to Cash Drawer
            newLedRow.Range(10).Value = 0
            newLedRow.Range(11).Value = txnID
            newLedRow.Range(12).Value = notes
            
            ' Movement 2: Wallet transferred to customer
            Set newLedRow = tblLed.ListRows.Add
            movCount = movCount + 1
            newLedRow.Range(1).Value = "MOV-" & Format(movCount, "000000")
            newLedRow.Range(2).Value = txnID
            newLedRow.Range(3).Value = txnDate
            newLedRow.Range(4).Value = txnTime
            newLedRow.Range(5).Value = dstAcc
            newLedRow.Range(6).Value = GetAccountName(dstAcc)
            newLedRow.Range(7).Value = "تحويل محفظة (صادر)"
            newLedRow.Range(8).Value = "تحويل رصيد للعميل " & custName & " - " & custPhone
            newLedRow.Range(9).Value = 0
            newLedRow.Range(10).Value = pAmount ' Outgoing from Wallet
            newLedRow.Range(11).Value = txnID
            newLedRow.Range(12).Value = notes
            
        Case "محفظة إلى كاش"
            ' Movement 1: Wallet received from customer
            Set newLedRow = tblLed.ListRows.Add
            movCount = movCount + 1
            newLedRow.Range(1).Value = "MOV-" & Format(movCount, "000000")
            newLedRow.Range(2).Value = txnID
            newLedRow.Range(3).Value = txnDate
            newLedRow.Range(4).Value = txnTime
            newLedRow.Range(5).Value = srcAcc
            newLedRow.Range(6).Value = GetAccountName(srcAcc)
            newLedRow.Range(7).Value = "تحويل محفظة (وارد)"
            newLedRow.Range(8).Value = "استلام تحويل محفظة من العميل " & custName
            newLedRow.Range(9).Value = rAmount ' Incoming to Wallet
            newLedRow.Range(10).Value = 0
            newLedRow.Range(11).Value = txnID
            newLedRow.Range(12).Value = notes
            
            ' Movement 2: Cash paid out to customer
            Set newLedRow = tblLed.ListRows.Add
            movCount = movCount + 1
            newLedRow.Range(1).Value = "MOV-" & Format(movCount, "000000")
            newLedRow.Range(2).Value = txnID
            newLedRow.Range(3).Value = txnDate
            newLedRow.Range(4).Value = txnTime
            newLedRow.Range(5).Value = dstAcc
            newLedRow.Range(6).Value = GetAccountName(dstAcc)
            newLedRow.Range(7).Value = "صرف نقدي (عميل)"
            newLedRow.Range(8).Value = "تسليم نقدية كاش للعميل " & custName
            newLedRow.Range(9).Value = 0
            newLedRow.Range(10).Value = dAmount ' Outgoing from Cash Drawer
            newLedRow.Range(11).Value = txnID
            newLedRow.Range(12).Value = notes
            
        Case Else ' Internal Transfer (Wallet to Wallet, Bank to Wallet, etc.)
            ' Movement 1: Source Account Outgoing
            Set newLedRow = tblLed.ListRows.Add
            movCount = movCount + 1
            newLedRow.Range(1).Value = "MOV-" & Format(movCount, "000000")
            newLedRow.Range(2).Value = txnID
            newLedRow.Range(3).Value = txnDate
            newLedRow.Range(4).Value = txnTime
            newLedRow.Range(5).Value = srcAcc
            newLedRow.Range(6).Value = GetAccountName(srcAcc)
            newLedRow.Range(7).Value = "تحويل داخلي صادر"
            newLedRow.Range(8).Value = "تحويل سيولة إلى " & dstAcc
            newLedRow.Range(9).Value = 0
            newLedRow.Range(10).Value = pAmount
            newLedRow.Range(11).Value = txnID
            newLedRow.Range(12).Value = notes
            
            ' Movement 2: Destination Account Incoming
            Set newLedRow = tblLed.ListRows.Add
            movCount = movCount + 1
            newLedRow.Range(1).Value = "MOV-" & Format(movCount, "000000")
            newLedRow.Range(2).Value = txnID
            newLedRow.Range(3).Value = txnDate
            newLedRow.Range(4).Value = txnTime
            newLedRow.Range(5).Value = dstAcc
            newLedRow.Range(6).Value = GetAccountName(dstAcc)
            newLedRow.Range(7).Value = "تحويل داخلي وارد"
            newLedRow.Range(8).Value = "استلام سيولة من " & srcAcc
            newLedRow.Range(9).Value = pAmount
            newLedRow.Range(10).Value = 0
            newLedRow.Range(11).Value = txnID
            newLedRow.Range(12).Value = notes
    End Select
    
    ' Movement 3 (If External Fee > 0): Deduct fee from specified fee account
    If extFee > 0 Then
        Set newLedRow = tblLed.ListRows.Add
        movCount = movCount + 1
        newLedRow.Range(1).Value = "MOV-" & Format(movCount, "000000")
        newLedRow.Range(2).Value = txnID
        newLedRow.Range(3).Value = txnDate
        newLedRow.Range(4).Value = txnTime
        newLedRow.Range(5).Value = feeAcc
        newLedRow.Range(6).Value = GetAccountName(feeAcc)
        newLedRow.Range(7).Value = "رسوم مزود خارجي"
        newLedRow.Range(8).Value = "رسوم خدمة مدفوعة للمزود لعملية " & txnID
        newLedRow.Range(9).Value = 0
        newLedRow.Range(10).Value = extFee
        newLedRow.Range(11).Value = txnID
        newLedRow.Range(12).Value = "رسوم شبكة / مزود"
    End If
    
    ' 4. RECALCULATE WORKBOOK
    ThisWorkbook.Calculate
    Application.ScreenUpdating = True
    
    ' 5. CLEAR FORM AND PREPARE NEXT ID
    Call ClearTransactionForm
    
    MsgBox "تم تسجيل العملية بنجاح!" & vbCrLf & _
           "رقم العملية: " & txnID & vbCrLf & _
           "صافي ربح العملية: " & Format(netProf, "#,##0.00") & " ج.م" & vbCrLf & _
           "تم ترحيل قيود دفتر الحركات وتحديث الأرصدة تلقائياً.", vbInformation, "نجاح التسجيل"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "حدث خطأ أثناء تسجيل العملية: " & Err.Description, vbCritical, "خطأ في النظام"
End Sub

Public Sub ClearTransactionForm()
    Dim wsEntry As Worksheet
    Set wsEntry = ThisWorkbook.Sheets("إدخال عملية")
    
    Dim nextNum As Long
    nextNum = Int(Rnd() * 9000 + 1000)
    wsEntry.Range("D4").Value = "TXN-" & Format(Date, "yyyymmdd") & "-" & Format(nextNum, "0000")
    wsEntry.Range("D5").Formula = "=TODAY()"
    wsEntry.Range("D6").Formula = "=TEXT(NOW(), ""hh:mm"")"
    wsEntry.Range("D7").Value = "كاش إلى محفظة"
    wsEntry.Range("D8").Value = ""
    wsEntry.Range("D9").Value = ""
    
    wsEntry.Range("G4").Value = "ACC-CASH-01"
    wsEntry.Range("G5").Value = "ACC-VF-01"
    wsEntry.Range("G6").Value = 0
    wsEntry.Range("G7").Value = 0
    wsEntry.Range("G8").Value = 0
    
    wsEntry.Range("J4").Value = 0
    wsEntry.Range("J5").Value = 0
    wsEntry.Range("J6").Value = "ACC-VF-01"
    wsEntry.Range("J7").Value = 0
    wsEntry.Range("J8").Formula = "=J4-J5-J7"
    
    wsEntry.Range("D11").Value = ""
    wsEntry.Range("G11").Value = "ناجحة"
End Sub

Public Sub ReverseTransactionUI()
    Dim targetTxnID As String
    targetTxnID = InputBox("أدخل رقم العملية التي ترغب في عكسها / إلغائها:", "عكس عملية سابقة")
    If Trim(targetTxnID) = "" Then Exit Sub
    
    Call ExecuteReversal(Trim(targetTxnID))
End Sub

Public Sub ExecuteReversal(ByVal targetID As String)
    On Error GoTo ErrRev
    Dim wsTxns As Worksheet, wsLedger As Worksheet
    Dim tblTxn As ListObject, tblLed As ListObject
    Dim findCell As Range, rIndex As Long
    
    Set wsTxns = ThisWorkbook.Sheets("سجل العمليات")
    Set wsLedger = ThisWorkbook.Sheets("دفتر الحركات")
    Set tblTxn = wsTxns.ListObjects("tbl_Transactions")
    Set tblLed = wsLedger.ListObjects("tbl_Ledger")
    
    Set findCell = tblTxn.ListColumns(1).DataBodyRange.Find(What:=targetID, LookIn:=xlValues, LookAt:=xlWhole)
    If findCell Is Nothing Then
        MsgBox "لم يتم العثور على العملية رقم (" & targetID & ")!", vbExclamation, "غير موجودة"
        Exit Sub
    End If
    
    rIndex = findCell.Row - tblTxn.HeaderRowRange.Row
    If tblTxn.ListColumns("Status").DataBodyRange.Cells(rIndex, 1).Value = "معكوسة" Then
        MsgBox "هذه العملية معكوسة بالفعل مسبقاً!", vbExclamation, "تنبيه"
        Exit Sub
    End If
    
    ' Mark as reversed
    tblTxn.ListColumns("Status").DataBodyRange.Cells(rIndex, 1).Value = "معكوسة"
    
    ' Find all ledger entries for this txn and post opposite movements
    Dim i As Long, movCount As Long
    Dim newLedRow As ListRow
    movCount = tblLed.ListRows.Count
    
    Application.ScreenUpdating = False
    For i = 1 To tblLed.ListRows.Count
        If tblLed.ListColumns(2).DataBodyRange.Cells(i, 1).Value = targetID Then
            Dim accID As String, accName As String, inAmt As Double, outAmt As Double
            accID = tblLed.ListColumns(5).DataBodyRange.Cells(i, 1).Value
            accName = tblLed.ListColumns(6).DataBodyRange.Cells(i, 1).Value
            inAmt = Val(tblLed.ListColumns(9).DataBodyRange.Cells(i, 1).Value)
            outAmt = Val(tblLed.ListColumns(10).DataBodyRange.Cells(i, 1).Value)
            
            Set newLedRow = tblLed.ListRows.Add
            movCount = movCount + 1
            newLedRow.Range(1).Value = "MOV-" & Format(movCount, "000000")
            newLedRow.Range(2).Value = targetID
            newLedRow.Range(3).Value = Format(Date, "yyyy-mm-dd")
            newLedRow.Range(4).Value = Format(Time, "hh:mm")
            newLedRow.Range(5).Value = accID
            newLedRow.Range(6).Value = accName
            newLedRow.Range(7).Value = "قيد عكسي"
            newLedRow.Range(8).Value = "إلغاء وعكس أثر الحركة التابعة للعملية " & targetID
            newLedRow.Range(9).Value = outAmt ' Swapped
            newLedRow.Range(10).Value = inAmt ' Swapped
            newLedRow.Range(11).Value = "REV-" & targetID
            newLedRow.Range(12).Value = "تم العكس آلياً لاستعادة الأرصدة"
        End If
    Next i
    
    ThisWorkbook.Calculate
    Application.ScreenUpdating = True
    
    MsgBox "تم إلغاء وعكس العملية (" & targetID & ") بنجاح!" & vbCrLf & _
           "تم توليد القيود العكسية واستعادة أرصدة الحسابات بدقة دون حذف السجل التاريخي.", vbInformation, "تم العكس بنجاح"
    Exit Sub
    
ErrRev:
    Application.ScreenUpdating = True
    MsgBox "حدث خطأ أثناء عكس العملية: " & Err.Description, vbCritical, "خطأ"
End Sub

Public Sub SaveDailyClosing()
    On Error GoTo ErrClose
    Dim wsClose As Worksheet
    Dim tblClose As ListObject, tblLog As ListObject
    Dim i As Long, logRow As ListRow, logCount As Long
    Dim hasDiff As Boolean
    
    Set wsClose = ThisWorkbook.Sheets("التقفيل اليومي")
    Set tblClose = wsClose.ListObjects("tbl_DailyClosing")
    Set tblLog = wsClose.ListObjects("tbl_ClosingLog")
    
    logCount = tblLog.ListRows.Count
    hasDiff = False
    
    Application.ScreenUpdating = False
    For i = 1 To tblClose.ListRows.Count
        Dim accID As String, accName As String, expBal As Double, actBal As Double, diff As Double
        Dim stat As String
        
        accID = tblClose.ListColumns(1).DataBodyRange.Cells(i, 1).Value
        accName = tblClose.ListColumns(2).DataBodyRange.Cells(i, 1).Value
        expBal = Val(tblClose.ListColumns(7).DataBodyRange.Cells(i, 1).Value)
        actBal = Val(tblClose.ListColumns(8).DataBodyRange.Cells(i, 1).Value)
        diff = actBal - expBal
        
        If diff = 0 Then
            stat = "متطابق"
        ElseIf diff < 0 Then
            stat = "عجز"
            hasDiff = True
        Else
            stat = "زيادة"
            hasDiff = True
        End If
        
        Set logRow = tblLog.ListRows.Add
        logCount = logCount + 1
        logRow.Range(1).Value = "LOG-" & Format(logCount, "0000")
        logRow.Range(2).Value = Format(Date, "yyyy-mm-dd")
        logRow.Range(3).Value = accID
        logRow.Range(4).Value = accName
        logRow.Range(5).Value = expBal
        logRow.Range(6).Value = actBal
        logRow.Range(7).Value = diff
        logRow.Range(8).Value = stat
        logRow.Range(9).Value = IIf(diff = 0, "مطابق تماماً", "يوجد فرق يحتاج مراجعة")
        logRow.Range(10).Value = "مسؤول الخزنة"
        logRow.Range(11).Value = "معتمد"
    Next i
    
    ThisWorkbook.Calculate
    Application.ScreenUpdating = True
    
    If hasDiff Then
        MsgBox "تم حفظ التقفيل اليومي في السجل بنجاح!" & vbCrLf & _
               "تنبيه: تم رصد فروق جرد (عجز / زيادة) في بعض الحسابات، يرجى مراجعة الأسباب وتوثيقها.", vbExclamation, "تقفيل به فروق"
    Else
        MsgBox "تم حفظ التقفيل اليومي بنجاح، وجميع الحسابات متطابقة تماماً 100%!", vbInformation, "تقفيل مطابق"
    End If
    Exit Sub
    
ErrClose:
    Application.ScreenUpdating = True
    MsgBox "حدث خطأ أثناء حفظ التقفيل: " & Err.Description, vbCritical, "خطأ"
End Sub

Public Sub RecordClosingDifference()
    Dim accID As String, diffAmt As String, diffReason As String
    accID = InputBox("أدخل كود الحساب الذي ترغب في تسجيل تسوية له (مثال: ACC-CASH-01):", "تسوية فرق جرد")
    If Trim(accID) = "" Then Exit Sub
    
    diffAmt = InputBox("أدخل مبلغ الفرق (موجب للزيادة، سالب للعجز):", "مبلغ الفرق")
    If Not IsNumeric(diffAmt) Then
        MsgBox "المبلغ غير صحيح!", vbCritical, "خطأ"
        Exit Sub
    End If
    
    diffReason = InputBox("أدخل سبب الفرق المعتمد:", "سبب التسوية")
    If Trim(diffReason) = "" Then diffReason = "تسوية جرد معتمدة"
    
    Dim wsLed As Worksheet, tblLed As ListObject, newRow As ListRow
    Set wsLed = ThisWorkbook.Sheets("دفتر الحركات")
    Set tblLed = wsLed.ListObjects("tbl_Ledger")
    
    Dim dVal As Double
    dVal = Val(diffAmt)
    
    Set newRow = tblLed.ListRows.Add
    Dim movCount As Long
    movCount = tblLed.ListRows.Count
    newRow.Range(1).Value = "MOV-" & Format(movCount, "000000")
    newRow.Range(2).Value = "ADJ-" & Format(Date, "yyyymmdd")
    newRow.Range(3).Value = Format(Date, "yyyy-mm-dd")
    newRow.Range(4).Value = Format(Time, "hh:mm")
    newRow.Range(5).Value = accID
    newRow.Range(6).Value = GetAccountName(accID)
    
    If dVal > 0 Then
        newRow.Range(7).Value = "تسوية زيادة جرد"
        newRow.Range(8).Value = "قيد تسوية زيادة معتمدة: " & diffReason
        newRow.Range(9).Value = dVal
        newRow.Range(10).Value = 0
    Else
        newRow.Range(7).Value = "تسوية عجز جرد"
        newRow.Range(8).Value = "قيد تسوية عجز معتمد: " & diffReason
        newRow.Range(9).Value = 0
        newRow.Range(10).Value = Abs(dVal)
    End If
    newRow.Range(11).Value = "ADJ-REF"
    newRow.Range(12).Value = diffReason
    
    ThisWorkbook.Calculate
    MsgBox "تم تسجيل قيد التسوية المحاسبي بنجاح!", vbInformation, "نجاح التسوية"
End Sub

Public Sub RefreshAll()
    ThisWorkbook.Calculate
    MsgBox "تم تحديث كافة المؤشرات والأرصدة والتقارير بنجاح!", vbInformation, "تحديث البيانات"
End Sub

' Fast Sheet Navigation
Public Sub Nav_Dashboard()
    ThisWorkbook.Sheets("لوحة التحكم").Activate
End Sub
Public Sub Nav_Entry()
    ThisWorkbook.Sheets("إدخال عملية").Activate
End Sub
Public Sub Nav_Transactions()
    ThisWorkbook.Sheets("سجل العمليات").Activate
End Sub
Public Sub Nav_Ledger()
    ThisWorkbook.Sheets("دفتر الحركات").Activate
End Sub
Public Sub Nav_Accounts()
    ThisWorkbook.Sheets("إدارة الحسابات").Activate
End Sub
Public Sub Nav_Loans()
    ThisWorkbook.Sheets("إدارة السلف").Activate
End Sub
Public Sub Nav_Custody()
    ThisWorkbook.Sheets("إدارة العهد").Activate
End Sub
Public Sub Nav_Closing()
    ThisWorkbook.Sheets("التقفيل اليومي").Activate
End Sub
Public Sub Nav_Reports()
    ThisWorkbook.Sheets("التقارير").Activate
End Sub
Public Sub Nav_Settings()
    ThisWorkbook.Sheets("الإعدادات").Activate
End Sub
Public Sub Nav_Testing()
    ThisWorkbook.Sheets("الاختبارات").Activate
End Sub

Private Function GetAccountName(ByVal accCode As String) As String
    On Error Resume Next
    Dim wsAcc As Worksheet, tblAcc As ListObject, fCell As Range
    Set wsAcc = ThisWorkbook.Sheets("إدارة الحسابات")
    Set tblAcc = wsAcc.ListObjects("tbl_Accounts")
    Set fCell = tblAcc.ListColumns(1).DataBodyRange.Find(What:=accCode, LookIn:=xlValues, LookAt:=xlWhole)
    If Not fCell Is Nothing Then
        GetAccountName = fCell.Offset(0, 1).Value
    Else
        GetAccountName = accCode
    End If
End Function
'''

def build_advanced_system():
    project_dir = r"c:\Users\mohamed\Desktop\xlsx project"
    base_file = os.path.join(project_dir, "Cash_Transfer_System_Base.xlsx")
    xlsm_file = os.path.join(project_dir, "Cash_Transfer_System.xlsm")
    xlsx_final = os.path.join(project_dir, "Cash_Transfer_System.xlsx")
    
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        print("Opening base workbook in Excel...")
        wb = excel.Workbooks.Open(base_file)
        
        # 1. Inject VBA Code Module
        print("Injecting VBA Automation Module...")
        vb_proj = wb.VBProject
        vb_comp = vb_proj.VBComponents.Add(1) # vbext_ct_StdModule
        vb_comp.Name = "mod_FinancialSystem"
        vb_comp.CodeModule.AddFromString(VBA_CODE)
        print("VBA Module injected successfully!")
        
        # 2. Add Form Buttons
        print("Adding interactive action buttons...")
        
        # Helper to add button shape
        # Excel msoShapeRoundedRectangle = 5
        def add_button(ws, left, top, width, height, text, macro, bg_rgb):
            shape = ws.Shapes.AddShape(5, left, top, width, height)
            shape.OnAction = macro
            shape.TextFrame2.TextRange.Characters.Text = text
            shape.TextFrame2.TextRange.Font.Name = "Segoe UI"
            shape.TextFrame2.TextRange.Font.Size = 11
            shape.TextFrame2.TextRange.Font.Bold = True
            shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF # White
            shape.Fill.ForeColor.RGB = bg_rgb
            shape.Line.Visible = False
            return shape
            
        # Sheet: إدخال عملية (Transaction Entry)
        ws_entry = wb.Sheets("إدخال عملية")
        # Green button for Save (RGB: 30, 126, 52 -> 0x347E1E in Win32 RGB: 52*65536 + 126*256 + 30)
        # Note: in win32 BGR vs RGB: RGB(r,g,b) = r + g*256 + b*65536
        rgb_green = 30 + (126 * 256) + (52 * 65536)
        rgb_slate = 73 + (80 * 256) + (87 * 65536)
        rgb_crimson = 200 + (35 * 256) + (51 * 65536)
        rgb_navy = 15 + (37 * 256) + (55 * 65536)
        
        # Add buttons above recent table (around row 13)
        add_button(ws_entry, 150, 240, 160, 32, "تسجيل وحفظ العملية", "SaveTransaction", rgb_green)
        add_button(ws_entry, 330, 240, 140, 32, "تفريغ الحقول", "ClearTransactionForm", rgb_slate)
        add_button(ws_entry, 490, 240, 160, 32, "إلغاء وعكس عملية", "ReverseTransactionUI", rgb_crimson)
        
        # Sheet: التقفيل اليومي (Daily Closing)
        ws_closing = wb.Sheets("التقفيل اليومي")
        add_button(ws_closing, 150, 480, 220, 32, "حفظ التقفيل اليومي بالأرشيف", "SaveDailyClosing", rgb_navy)
        add_button(ws_closing, 390, 480, 200, 32, "تسجيل تسوية عجز أو زيادة", "RecordClosingDifference", rgb_crimson)
        
        # Sheet: لوحة التحكم (Dashboard)
        ws_dash = wb.Sheets("لوحة التحكم")
        # Navigation Bar Buttons at row 22
        nav_items = [
            ("تسجيل عملية", "Nav_Entry"),
            ("سجل العمليات", "Nav_Transactions"),
            ("دفتر الحركات", "Nav_Ledger"),
            ("إدارة الحسابات", "Nav_Accounts"),
            ("السلف", "Nav_Loans"),
            ("العهد", "Nav_Custody"),
            ("التقفيل اليومي", "Nav_Closing"),
            ("التقارير", "Nav_Reports"),
            ("تحديث المؤشرات", "RefreshAll"),
        ]
        nav_left = 50
        for label, mac in nav_items:
            add_button(ws_dash, nav_left, 490, 100, 28, label, mac, rgb_navy if mac != "RefreshAll" else rgb_green)
            nav_left += 110
            
        # 3. Add Native Excel Charts on Dashboard
        print("Generating Dashboard Charts...")
        try:
            # Chart 1: Donut Chart for Liquidity Distribution (from I13:J17)
            chart_obj1 = ws_dash.ChartObjects().Add(450, 230, 350, 240)
            chart1 = chart_obj1.Chart
            chart1.ChartType = 80 # xlDoughnut
            chart1.SetSourceData(ws_dash.Range("I13:J17"))
            chart1.HasTitle = True
            chart1.ChartTitle.Text = "هيكل توزيع السيولة النقدية ورأس المال"
            chart1.ChartTitle.Font.Name = "Segoe UI"
            chart1.ChartTitle.Font.Size = 10
            chart1.ApplyDataLabels()
            
            # Chart 2: Clustered Bar Chart for Wallet Lines (B14:C21, F14:F21)
            chart_obj2 = ws_dash.ChartObjects().Add(50, 230, 380, 240)
            chart2 = chart_obj2.Chart
            chart2.ChartType = 57 # xlBarClustered
            # Set source: line names in C14:C21 and balances in F14:F21
            chart2.SetSourceData(ws_dash.Range("C14:C21,F14:F21"))
            chart2.HasTitle = True
            chart2.ChartTitle.Text = "أرصدة خطوط فودافون وأورنج الرئيسية"
            chart2.ChartTitle.Font.Name = "Segoe UI"
            chart2.ChartTitle.Font.Size = 10
            chart2.HasLegend = False
            print("Dashboard Charts created successfully!")
        except Exception as ce:
            print("Chart creation notice:", ce)
            
        # 4. Save as .XLSM (FileFormat 52)
        print(f"Saving workbook as XLSM: {xlsm_file}")
        wb.SaveAs(xlsm_file, 52) # 52 = xlOpenXMLWorkbookMacroEnabled
        
        # 5. Also save copy as .XLSX (FileFormat 51)
        print(f"Saving workbook as XLSX: {xlsx_final}")
        wb.SaveCopyAs(xlsx_final)
        
        wb.Close(False)
        print("Build completed successfully!")
    except Exception as e:
        print("Error during advanced build:", e)
        raise e
    finally:
        excel.Quit()

if __name__ == "__main__":
    build_advanced_system()

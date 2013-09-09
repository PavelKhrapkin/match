Attribute VB_Name = "REPORT_ADSK"
'---------------------------------------------------------------------------------
' REPORT_ADSK  - √енераци€ внешних отчетов по продуктам Autodesk
'
' - PaidADSK(PayK, Acc, Dat, Spec,Sbs) - анализ спецификации Spec, ее св€зь с Autodesk
' - ContrADSKbySN(SN)           - возвращает  онтракт Autodesk по SN в SF
' - ADSKqty(Acc, Desck, Dat, Contr) - сколько у ќрганизации Acc мест типа Descr
' T testDIC_GoodADSK()          - отладка IsGoodInSpec и FindInLst
' - IsGoodInSpec(Good, Spec)    - распознает есть ли товар Good в Spec
' - IsContrADSKinSF(ContrADSK)  - возвращает TRUE если контракт ContrADSK есть в SF
'
'   22.8.2012

Option Explicit

Sub PaidADSKqtys()
'
' S PaidADSKqtys() - получение статистики по Seats ADSK из ѕлатежей 1—
'   22/8/12
    
    Dim P As TOCmatch
    
    P = GetRep(PAY_SHEET)
    
    SheetInit ADSKstatistics
    
    Dim i, j, Qty As Integer, Good As String, Descr As String
    Dim Sbs As Boolean, Consulting As Boolean
    Dim Dat As Date

    With Sheets(PAY_SHEET)
        For i = 2 To P.EOL
            Progress i / P.EOL
            If Trim$(.Cells(i, PAYDOC_COL)) <> "" Then
                Good = .Cells(i, PAYGOOD_COL)
                Dat = .Cells(i, PAYDATE_COL)
                If GoodType(Good) = WE_GOODS_ADSK Then
                    For j = 0 To 999
                        Descr = ADSK_SpecItem(Good, j, Sbs, Consulting, Qty)
                        If Descr = "" Then Exit For
                        If Sbs Then
                            PutInTab "ADSK_Subs", Qty, Descr, Dat
                        Else
                            If Not Consulting Then PutInTab "ADSK_Lic", Qty, Descr, Dat
                        End If
                    Next j
                End If
            End If
        Next i
    End With
End Sub
Sub SheetInit(SheetName)
'
' - SheetInit(SheetName) - инициирует лист SheetName по форме в именованном Range
'   22.8.12
    Dim Form0 As String
    Form0 = "FORM_" & SheetName
   
    With DB_MATCH
        On Error Resume Next
        .Sheets(SheetName).Delete
        .Sheets.Add After:=.Sheets(.Sheets.count)
        .Sheets(.Sheets.count).Name = SheetName
        On Error GoTo 0
        With .Sheets(SheetName)
            .Tab.Color = Range(Form0).Cells(1, 1).Interior.Color
            .Activate
            For i = 1 To Range(Form0).Columns.count
                Range(Form0).Columns(i).Copy Destination:=.Cells(1, i)
                .Columns(i).ColumnWidth = .Cells(1, i)
            Next i
            .Rows(1).Delete
        End With
    End With
End Sub
Sub PutInTab(Rng, Val, Cat, Dat)
'
'- PutInTab(Rng, Val, Cat, Dat) - заносит в Range Rng Val в поле по Cat и Dat
'   15/5/12

    Dim i As Integer
    Dim Str As Range
    Dim h As Range
    Dim HDR_Date As Date
    Dim y, Y0, M, M0
    
    If Not IsDate(Dat) Then GoTo Error
    Y0 = Year(Dat)
    M0 = Month(Dat)
       
    For Each Str In Range(Rng).Rows
        If Str.Cells(1, 1) = Cat Then
            i = 1
            For Each h In Range(Rng).Columns
                If IsDate(h.Cells(1, 1)) Then
                    y = Year(h.Cells(1, 1))
                    M = Month(h.Cells(1, 1))
                    If y = Y0 And M = M0 Then
                        Str.Cells(1, i) = Str.Cells(1, i) + Val
                        Exit Sub
                    End If
                End If
                i = i + 1
            Next h
        End If
    Next Str

Error: MsgBox Cat & " или " & Dat & " не найден!", , "ERROR"
    Stop
End Sub

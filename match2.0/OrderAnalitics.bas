Attribute VB_Name = "OrderAnalitics"
'---------------------------------------------------------------------------------
' OrderAnalitics  - ������ �� ����� ������� �.�������
'
' S   GetInv1C(InvCol,PayN_Col, _       - ������� ����� ������ � �������� PayN �� ������,
'         StrInvCol,DateCol,[Str2Inv])    ��������� �� �������, ���������� ���� � �� ����.
' [*] OrderPass()     - ������ �� ����� ������� Orders
'
'   4.9.2012

Option Explicit
Sub GetInv1C(InvCol As Integer, PayN_Col As Integer, _
    StrInvCol As Integer, DateCol As Integer, Optional Str2InvCol As Integer = 0)
'
'S GetInv1C(InvCol,PayN_Col, _       - ������� ����� ������ � �������� �� ������ 1�,
'      StrInvCol,DateCol,[Str2Inv])    ��������� �� �������, ���������� ���� � �� ����.
' ----- ���������, ������������ � ������� ��������� -------
' 1.InvCol      - ����� ������� � MyCol, ���� ����������� ��������� ������ - ���� 1�
' 2.PayN_Col    - ����� ������� � MyCol, ���� �������� ��������� ����� ������ � �������� 1�
' 3.StrInvCol   - ����� ������� � ������� - "����� ����� 1�"
' 4.DateCol     - ����� ������� - �������� � ���� ����� 1�
' 5.[Str2InvCol]- �������������� ������� � ������� ����� 1�
'-----------------------------------------------------------
'           * ������������ ��� ��� ��� �������� ������ "������" ��� "�����"
'           * ������� Str2InvCol - ��������� �������������� ������ �� ������
'   26.8.12
' 4.9.12 - StepIn

    StepIn
    
    Dim DocTo As String ' ��� �������� ��������� - ������
    Dim RepSF As TOCmatch, RepSForder As TOCmatch, RepP As TOCmatch, RepTo As TOCmatch
    Dim Str As String, Inv As String, Str2 As String, Inv2 As String, D As Date
    Dim i As Integer, i1C As Integer
    Dim X
    
    DocTo = ActiveSheet.Name
    RepTo = GetRep(DocTo)
    RepP = GetRep(PAY_SHEET)
    EOL_PaySheet = RepP.EOL
    

    With Workbooks(RepTo.RepFile).Sheets(RepTo.SheetN)
        .Activate
        For i = 2 To RepTo.EOL
            Progress i / RepTo.EOL
            
            Str = SeekInv(.Cells(i, StrInvCol))
            
            X = .Cells(i, DateCol): D = "0:0"
            If IsDate(X) Then D = X
            
            Str2 = ""
            If Str2InvCol > 0 Then Str2 = SeekInv(.Cells(i, Str2InvCol))
            
            If Str = Str2 Then
            ElseIf Str = "" Or Str2 = "" Then
                Str = Str & Str2
            Else
                Str = Str & "+" & Str2
            End If
            .Cells(i, InvCol) = Str
            
            If IsInv1C(Str, D, i1C) Then .Cells(i, PayN_Col) = i1C
        Next i
    End With

End Sub

Sub OrderPass()
'
' [*] OrderPass()   - ������ �� ����� �������, ������������ ����� �������
'   28.4.12
'   18.6.12 - ���������, 5 ��������� �������

'------ INITIALIZATION AND LOCAL DECLARATION SECTION ---------------------
    Dim Inv1C As String     '= ����������� �� ����� ���� 1�
    Dim Dat As String       '���� ���� ����� CSD
    Dim PaidDat As String   '���� ���� ������ �������
    Dim Sale1C As String    '���� �������� � 1�
    Dim Client As String    '���� ������ 1�
    Dim iOL As Integer      '= ����� ������ � �������
    Dim i1� As Integer      '= ����� ������ � ��������
    
    Dim i As Integer
    
    EOL_OrderList = ModStart(OrderList, "������ �� �������: ��������� � SF", True)
    
    CheckSheet OrderList, 1, OL_ORDERN_COL, OrderListStamp
    ClearSheet NewOrderList, Range("HDR_NewOrderList")
    
    EOL_SForders = EOL(SForders)
    
'---------------------- CODE SECTION -----------------------------------
    With Sheets(OrderList)
        '-- ��������� EOL_OrderList �������� �����
        EOL_OrderList = EOL_OrderList - OL_MIN_RESLINES
        Do While .Cells(EOL_OrderList, OL_ORDERN_COL) = ""
            EOL_OrderList = EOL_OrderList - 1
        Loop
         
        For i = 2 To EOL_OrderList
            Progress i / EOL_OrderList
            
If i >= 113 Then
i = i
End If
            Inv1C = .Cells(i, OL_INV_1C_COL)
            If Trim(Inv1C) = "" Then
                Inv1C = SeekInv(.Cells(i, OL_ORDERN_COL))
            End If
            Dat = .Cells(i, OL_CSDINVDAT_COL)
            If IsInv1C(Inv1C, Dat, i1�) Then
                With Sheets(PAY_SHEET)
                    PaidDat = .Cells(i1�, PAYDATE_COL)
                    Inv1C = .Cells(i1�, PAYINVOICE_COL)
                    Sale1C = .Cells(i1�, PAYSALE_COL)
                    Client = .Cells(i1�, PAYACC_COL)
                    
                End With
            Else
                PaidDat = "": Inv1C = "": Client = "": Sale1C = ""
            End If
            .Cells(i, OL_PAIDDAT_COL) = PaidDat
            .Cells(i, OL_INV1C_COL) = Inv1C
            .Cells(i, OL_SALE1C_COL) = Sale1C
            .Cells(i, OL_ACC1C_COL) = Client
            
'            OrderN = .Cells(i, OL_ORDERN_COL)
'            If Not IsOrderN(OrderN, iOL) Then
'                NewOrder (i)
'            End If
        Next i
    End With
'----------------------- SUMMARY SECTION -------------------------------
    ModEnd OrderList
End Sub
Function IsInv1C(Str, Dat, i1C) As Boolean
'
' - IsInv1C(Str, Dat, i1C)  - ���������� TRUE � ����� ������ � �������� 1�,
'                             ���� ���� �� Str ��������� � ������
'   18.6.12
    
    Const PO_DAYS = 50      ' ���������� ����� ���� �� ������� �� ������
    
    Dim Inv1C As String     '���� "����" ������� 1�
    Dim D As Date           '���� "���� ������� �����"
    Dim D_Min As Date
    Dim D_Max As Date
    
    IsInv1C = False
    If Not IsDate(Dat) Or Str = "" Then Exit Function
    
    D_Min = CDate(Dat) - PO_DAYS
    D_Max = CDate(Dat) + PO_DAYS
    
    For i1C = 2 To EOL_PaySheet
        Inv1C = DB_1C.Sheets(PAY_SHEET).Cells(i1C, PAYINVOICE_COL)
        If InStr(Inv1C, Str) <> 0 Then
            On Error Resume Next
            D = DB_1C.Sheets(PAY_SHEET).Cells(i1C, PAYDATE_COL)
            On Error GoTo 0
            If D < D_Max And D > D_Min Then
                IsInv1C = True
                Exit Function
            End If
        End If
    Next i1C
End Function
Sub testInv1C()
    Dim i1C As Integer
    ModStart PAY_SHEET, "���� Inv1C"
    Call IsInv1C("��-278", "01.06.12", i1C)
End Sub

Function IsOrderN(OrderN, iOL) As Boolean
'
' ���� ����� OrderN ���� � SF, ���������� TRUE
' 28.4.12

    IsOrderN = False
    Dim i As Integer
    For i = 2 To EOL_SForders
        If OrderN = Sheets(SForders).Cells(i, SFORDERS_ORDERN_COL) Then
            IsOrderN = True
            Exit Function
        End If
    Next i
    Exit Function
End Function
Sub WrOrderMyCol()
'
' S WrOrderMyCol()
' 4.9.12

    StepIn
    
    Dim R As TOCmatch   ' ������� �������� "������"
    Dim P As Range      ' ����� � ����������
    Dim i As Long, OrderLine As Long, Ind As Long
    Dim X As String, Y As String, IsErr As Boolean
    
    R = GetRep("������")
    Set P = DB_STOCK.Sheets("Forms").Range("Orders_MyCol")
    
    With DB_STOCK.Sheets(R.SheetN)
        For OrderLine = 2 To R.EOL
            Progress OrderLine / R.EOL
            For i = 1 To R.MyCol
                Ind = P.Cells(4, i)
                If Ind > 0 Then
                    X = .Cells(OrderLine, Ind)
                    Y = Adapter(P.Cells(5, i), X, P.Cells(6, i), IsErr)
                    If IsErr Then
                        .Cells(OrderLine, i) = ""
                        .Cells(OrderLine, i).Interior.Color = rgbRed
                    Else
                        .Cells(OrderLine, i) = Y
                    End If
                End If
            Next i
        Next OrderLine
    End With
End Sub
Sub NewOrd()
'
' S NewOrder - ������ ����� ������� � ���� NewOrderList ��� �������� � SF
'   5.9.12

    StepIn

    Dim Ord As TOCmatch '������
    Dim i As Long
    
    Ord = GetRep(OrderList)
    
    With DB_MATCH.Sheets(NewOrderList)
        For i = 2 To Ord.EOL
            Progress i / Ord.EOL
            If .Cells(i, OL_IDSFORDER_COL) = "" And .Cells(i, OL_IDSF_COL) <> "" Then
                WrNewSheet NewOrderList, DB_STOCK.Sheets(OrderList), i
            End If
        Next i
    End With
End Sub

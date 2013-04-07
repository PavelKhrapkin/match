Attribute VB_Name = "BTOmailHandling"
'------------------------------------------------------------------------------------
' BTOhandling - ��������� e-mail'�� �� CSD �� �������� ���������� SN �� �����
'
'   6.4.13
'
'[*] BTO_Mail_track()   - ������ � ��������� ����� BTOmails
' -  BTOmailHandle (SN, BTOmsg, BTOmsgLines) - ��������� ������ ���
' -  IsSNonStock(SN)    - return TRUE if SN is registered on Stock

Option Explicit

Sub BTO_Mail_track()
'
'[*] BTO_Mail_track() - ������ � ��������� ����� BTOmails
'
' When string in file contains BTOstamp, read mail - seek SN on Stock
'   12.6.12
'   15.6.12 - ������ ������ Autodesk �����������. ������������� ���,
'             ����� �� �������� SN
'   18.6.12 - �������� ����� CSD �� ����� �������
'   14.11.12 - ������� � match 2.0
'    6.04.13 - ��������� ����� TOCmatch � ��������� � Stock.xlsx, ��������� ���
'
'------ INITIALIZATION AND LOCAL DECLARATION SECTION ---------------------
'''''    Const BTOfileName = "BTOmails.txt"  ' ������� ���� - ������ �� Outlook
    Dim R As TOCmatch       '������� ���� BTOmails ��� ���������
    Dim iMail As Integer    '= ����� ����������� ������
'''    Dim iSN As Integer      '= ����� SN, �� ����������� �� ������
'''    Dim iADSK As Integer    '= ����� ������ ADSK �� ����� BTOmails.txt
    Dim MailDate As Date    '= ���� � ����� ������ ��� � mail Subject
    Dim CSD_Inv As String   '= �����/���� CSD
    Dim Descr As String     '�������� �������� Autodesk
    Dim SN As String        '= ������� SN
    Dim Seats As Long       '���������� ����
    Dim i As Long, S As String
    
    StepIn
    R = GetRep(BTOmails)
    
'---------------------- CODE SECTION -----------------------------------
    With Workbooks(R.RepFile).Sheets(R.SheetN)
        For i = 1 To R.EOL
            Progress i / R.EOL
            S = .Cells(i, 1)
            If InStr(S, "���: ���������� �� ��������") <> 0 Then
                iMail = iMail + 1
                MailDate = Mid(S, 2, WorksheetFunction.FindB("]", S) - 2)
            ElseIf InStr(S, "����:#") <> 0 Then
                CSD_Inv = Mid(S, 7)
            ElseIf InStr(S, "Auto") Then
'''                SN = RemIgnoredSN()
                BTOmailHandle MailDate, CSD_Inv, Descr, SN, Seats
            ElseIf InStr(S, "------") Then
                CSD_Inv = "": MailDate = 0
            End If
        Next i
    End With
'----------------------- SUMMARY SECTION -------------------------------
'''    Columns("A:J").Select           ' ����� ��� WrapText
'''    Selection.WrapText = False
'''
'''    MS "� ����� " & BTOmails & " ����������� " _
'''        & Str$(iMail) & " ����� ���, ��� " & Str$(iSN) _
'''        & " �� ��� �������� �� ������ �� �������"
''''''    ModEnd STOCK_SHEET
End Sub
Function BTOmailHandle(MailDate As Date, CSD_Inv As String, Descr As String, SN As String, Seats As Long) As Boolean
'''''Function BTOmailHandle(SN, BTOmsg, BTOmsgLines) As Boolean
'
' - BTOmailHandle (SN, BTOmsg, BTOmsgLines) - ��������� ������ ���
'       ���������� FALSE, ���� ������ ���������� �� ������� ���
'       ��������� � ��� SN ��� �������� �� ������ � ������������ �� ����
'   12.6.12
'   15.6.12 - ��������� ������� "�������� �� ������", "���� ������" � "���� 1�"

    Dim Sale As String      '���� BTO "��������"
    Dim Client As String    '���� BTO "��������"
    Dim Delivery As String  '���� "�������� �� ������"
    Dim Paid As String      '���� "���� ������ ����� � 1�"
    Dim Inv1C As String     '���� "���� 1�"
    Dim iStock As Integer   '= ����� ������ �� ��������� �����
    Dim iCSD As Integer     '= �� ����� �������
    Dim SN_SF As SNatr      '= ��������� SN � SF
    Dim iSF As Integer      '= ����� ������ � ������ ADSKfrSF �� SN
    Dim GoodADSK As String  '���� ��� "����� ADSK" - ������ �� ������
    
    Dim Msg As String
    Const MaxStrLen = 100
    Dim i As Integer
    
    Const InvCSD = 37   '������� ����� CSD � ������ Subject ����� ����
'    Dim iCSD As Integer '= ����� ������� - ����� CSD � Subject
    
    EOL_BTO = EOL_BTO + 1
    
'---- ��������� ���� � ������� ������, ��������� CSD �� Subject
    With Sheets(BTO_SHEET)
        For i = 2 To MaxStrLen
            If Mid(BTOmsg(1), i, 1) = "]" Then
                .Cells(EOL_BTO, BTO_DATE_COL) = Mid(BTOmsg(1), 2, i - 2)
                iCSD = i + InvCSD
                .Cells(EOL_BTO, BTO_CSDATR_COL) = _
                    Trim(Mid(BTOmsg(1), iCSD, Len(BTOmsg(1)) - iCSD - 7))
                Exit For
            End If
        Next i
        
'---- ������ � ������ ������� ������ ���
        Msg = ""
        For i = 1 To BTOmsgLines
            If Not InStr(BTOmsg(i), "-------") <> 0 Then
                Msg = Msg & BTOmsg(i) & vbCrLf
            End If
        Next i
        .Cells(EOL_BTO, BTO_MAIL_COL) = Msg

'---- ������ �� ������� "����� ADSK"
        For i = 2 To BTOmsgLines
            If InStr(BTOmsg(i), "Auto") <> 0 Then
                .Cells(EOL_BTO, BTO_GOOD_COL) = _
                    Mid(BTOmsg(i), 3, Len(BTOmsg(i)) - 3)
                Exit For
            End If
        Next i
'---- ������ � ������� ����� CSD
        If IsCSDinv(.Cells(EOL_BTO, BTO_CSDATR_COL), iCSD) Then
            With Sheets(OrderList)
                Paid = .Cells(iCSD, OL_PAIDDAT_COL)
                Inv1C = .Cells(iCSD, OL_INV1C_COL)
            End With
        Else
            Paid = "": Inv1C = ""
        End If
    
'---- ������ � SN
        .Cells(EOL_BTO, BTO_SN_COL) = SN
               
        If Len(SN) <> 12 Then
            Sale = "<-- ��� SN � ������ ��� -->"
            Client = "": Delivery = "": Sale = "": Inv1C = "": Paid = ""
        Else
            If IsSNonStock(SN, iStock) Then
               Delivery = Sheets(STOCK_SHEET).Cells(iStock, STOCK_DELIVERY_COL)
            End If
'---- SN �� SF
            SN_SF = SNinSFatr(SN, iSF)  '<<< ������� SN � SF >>>
            If SN_SF.ErrFlag Then
                Sale = "<-!- ����������� � SF -!->"
                Client = ""
                ErrMsg TYPE_ERR, "� SF ��� SN=" & SN
            Else
                
'---- ������ � BTOlog
                With Sheets(ADSKfrSF)
                    Sale = .Cells(iSF, SFADSK_SALE_COL)
                    Client = .Cells(iSF, SFADSK_ACC1C_COL)
                End With
            End If
        End If
        .Cells(EOL_BTO, BTO_DELIVERY_COL) = Delivery
        .Cells(EOL_BTO, BTO_PAID_DATE_COL) = Paid
        .Cells(EOL_BTO, BTO_INV_1C_COL) = Inv1C
        .Cells(EOL_BTO, BTO_SALE_COL) = Sale
        .Cells(EOL_BTO, BTO_CLIENT_COL) = Client
    End With
End Function
Function IsSNonStock(SN, iStock) As Boolean
'
' - IsSNonStock(SN, iStock)    - return TRUE if SN is registered on Stock
'   11.6.12
'   15.6.12 ���������� ����� ������ �� ������
    
    Dim i As Integer
    
    IsSNonStock = False
    If SN = "" Then Exit Function
    With Sheets(STOCK_SHEET)
        For i = 2 To EOL_Stock
            If InStr(.Cells(i, STOCK_SN_COL), SN) <> 0 Then
'                Client = .Cells(i, STOCK_CLIENT_COL)
'                Dat = .Cells(i, STOCK_DATE_COL)
                IsSNonStock = True
                iStock = i
                Exit Function
            End If
        Next i
    End With
End Function
Function IsCSDinv(Str, iCSD) As Boolean
'
' - IsCSDinv(Str, iCSD) - ���������� TRUE � ����� ������,
'              ���� ����� ������ ������ ����� ������� CSD
'   17.6.12

    Dim Inv As String   '���� "� ����� CSD" �������
    Dim MsgInv As String '= ����� ����� ��� � ������ ���
    Dim S() As String   '= ������ ��� ������� Str
    Dim Dat As Date     '���� "���� ����� CSD" �������
    Dim MsgDat As Date  '= "���� ����� CSD" �� ������ ���
    
    IsCSDinv = False
    If Str = "" Then Exit Function

    S = Split(Str, " ")
    If S(1) <> "��" Then Exit Function
    MsgInv = S(0)
    MsgDat = S(2)
    
    With Sheets(OrderList)
        For iCSD = 2 To EOL_OrderList
If iCSD >= 437 Then
S = S
End If
            Inv = .Cells(iCSD, OL_CSDINVN_COL)
            Dat = .Cells(iCSD, OL_CSDINVDAT_COL)
            If Inv = MsgInv And Abs(Dat - MsgDat) < 5 Then
                IsCSDinv = True
                Exit Function
            End If
        Next iCSD
    End With
End Function

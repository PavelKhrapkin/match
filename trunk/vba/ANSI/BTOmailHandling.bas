Attribute VB_Name = "BTOmailHandling"
'------------------------------------------------------------------------------------
' BTOhandling - ��������� e-mail'�� �� CSD �� �������� ���������� SN �� �����
'
'   14.6.12
'
'[*] BTO_Mail_track()   - ������ � ��������� ����� BTOmails
' -  BTOmailHandle (SN, BTOmsg, BTOmsgLines) - ��������� ������ ���
' -  IsSNonStock(SN)    - return TRUE if SN is registered on Stock

Option Explicit

Sub BTO_Mail_track()
'
'[*] BTO_Mail_track() - ������ � ��������� ����� BTOmails.txt
'
' When string in file contains BTOstamp, read mail - seek SN on Stock
'   12.6.12
'   14.6.12 - ������ ������ Autodesk �����������. ������������� ���,
'              ����� �� �������� SN

'------ INITIALIZATION AND LOCAL DECLARATION SECTION ---------------------
    Const BTOfileName = "BTOmails.txt"
    Dim iMail As Integer    '= ����� ����������� ������
    Dim iSN As Integer      '= ����� SN, �� ����������� �� ������
    Dim iADSK As Integer    '= ����� ������ ADSK �� ����� BTOmails.txt
    Dim SN As String        '= ������� SN
    Const MaxMailLines = 100
    Dim BTOmsg(1 To MaxMailLines) As String ' ������ �� ���
    Dim i, j
       
    iMail = 0: iSN = 0
    
    ModStart STOCK_SHEET, "��������� ����� ��� �� CSD", True
    CheckSheet STOCK_SHEET, 1, STOCK_PRODUCT_COL, STOCK_STAMP
    ClearSheet BTO_SHEET, Range("HDR_BTOmails")
    
    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\ADSK\"
    Open BTOfileName For Input As #1
    
'---------------------- CODE SECTION -----------------------------------
    i = 1
    Do While Not EOF(1)
        Progress i / 1865
        Line Input #1, BTOmsg(1)
        If InStr(BTOmsg(1), "���: ���������� �� ��������") <> 0 Then
            iMail = iMail + 1
            j = 1
            Do                  ' read BTO mail into BTOmsg array
                j = j + 1
                Line Input #1, BTOmsg(j)
                If InStr(BTOmsg(j), "Auto") <> 0 Then iADSK = j
                If InStr(BTOmsg(j), "--------") Then Exit Do
            Loop While j < MaxMailLines
            
            SN = RemIgnoredSN(BTOmsg(iADSK) & BTOmsg(iADSK + 1))
            
            If BTOmailHandle(SN, BTOmsg, j) Then iSN = iSN + 1
            
            i = i + j
        Else
            i = i + 1
        End If
    Loop
    Close #1
'----------------------- SUMMARY SECTION -------------------------------
    Columns("F:G").Select           ' ������ ��� ��� WrapText
    Selection.WrapText = False
    
    MS "� ����� " & BTOfileName & " ����������� " _
        & Str$(iMail) & " ����� ���, ��� " & Str$(iSN) _
        & " �� ��� �������� �� ������ �� �������"
    ModEnd STOCK_SHEET
End Sub
Function BTOmailHandle(SN, BTOmsg, BTOmsgLines) As Boolean
'
' - BTOmailHandle (SN, BTOmsg, BTOmsgLines) - ��������� ������ ���
'       ���������� FALSE, ���� ������ ���������� �� ������� ���
'       ��������� � ��� SN ��� �������� �� ������ � ������������ �� ����
'   12.6.12

    Dim Sale As String      '���� BTO "��������"
    Dim Client As String    '���� BTO "��������"
    Dim SN_SF As SNatr      '= ��������� SN � SF
    Dim iSF As Integer      '= ����� ������ � ������ ADSKfrSF �� SN
    Dim GoodADSK As String  '���� ��� "����� ADSK" - ������ �� ������
    
    Dim Msg As String
    Const MaxStrLen = 100
    Dim i As Integer
    
    Const InvCSD = 37   '������� ����� CSD � ������ Subject ����� ����
    Dim iCSD As Integer '= ����� ������� - ����� CSD � Subject
    
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
    
'---- ������ � SN
        .Cells(EOL_BTO, BTO_SN_COL) = SN
               
        If Len(SN) <> 12 Then
            Sale = "<-- ��� SN � ������ ��� -->"
            Client = ""
        ElseIf IsSNonStock(SN) Then
            Sale = "<-- SN �������� �� ������ -->"
            Client = ""
        Else
            SN_SF = SNinSFatr(SN, iSF)  '<<< ������� SN � SF >>>
            If SN_SF.ErrFlag Then
                Sale = "<-!- ����������� � SF -!->"
                Client = ""
                ErrMsg TYPE_ERR, "� SF ��� SN=" & SN
            Else
                With Sheets(ADSKfrSF)
                    Sale = .Cells(iSF, SFADSK_SALE_COL)
                    Client = .Cells(iSF, SFADSK_ACC1C_COL)
                End With
            End If
        End If
        .Cells(EOL_BTO, BTO_SALE_COL) = Sale
        .Cells(EOL_BTO, BTO_CLIENT_COL) = Client
    End With
End Function
Function IsSNonStock(SN) As Boolean
'
' - IsSNonStock(SN)    - return TRUE if SN is registered on Stock
'   11.6.12
    
    Dim i As Integer
    
    IsSNonStock = False
    If SN = "" Then Exit Function
    With Sheets(STOCK_SHEET)
        For i = 2 To EOL_Stock
            If InStr(.Cells(i, STOCK_SN_COL), SN) <> 0 Then
'                Client = .Cells(i, STOCK_CLIENT_COL)
'                Dat = .Cells(i, STOCK_DATE_COL)
                IsSNonStock = True
                Exit Function
            End If
        Next i
    End With
End Function


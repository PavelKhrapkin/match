Attribute VB_Name = "BTO_Mail_tracking"
Option Explicit

Sub BTO_Mail_trace()
'
' <*> BPO_Mail_trace()  - Create file BTOmails.txt from selected Mails in Folder "Заказы в CSD"
'
' Pavel Khrapkin 12.6.2012

    Dim myNameSpace As Outlook.NameSpace
    Dim myItems As Outlook.Items
    Dim Mail As MailItem
    Dim fld As Folder
    Dim i As Integer
    
    
    Dim F
    
    Set myNameSpace = Application.GetNamespace("MAPI")
    Set F = myNameSpace.Folders

    Set myItems = F.Item(2).Folders.Item(17).Items
    myItems.Sort "ReceivedTime", True
    
    ChDir "C:\Users\Пользователь\Desktop\Работа с Match\SFconstrTMP\ADSK\"
    Open "BTOmails.txt" For Append As #1
    
    For Each Mail In myItems
        If InStr(Mail.Subject, "Обновление по подписке") <> 0 Then
            Print #1, "[" & Mail.ReceivedTime & "]" & Mail.Subject & vbCrLf & Mail.Body _
                & vbCrLf & "-------------------------------------------------------"
        End If
    Next Mail
    Close #1
End Sub

Sub CompanyChange()
    Dim ContactsFolder As Folder
    Set ContactsFolder = Session.GetDefaultFolder(olFolderContacts)
    MsgBox ("Contacts Found: " & ContactsFolder.Items.Count)
End Sub
Sub t()
    ChDir "C:\Users\Пользователь\Desktop\Работа с Match\SFconstrTMP\ADSK\"

    Open "BTOtest.txt" For Append As #1
    Print #1, "Subj text " & Timer
    Print #1, "Subj text " & Timer
    Print #1, "Subj text " & Timer
End Sub

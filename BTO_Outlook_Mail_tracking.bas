Attribute VB_Name = "BTO_Outlook_Mail_tracking"
'------------------------------------------------------------------------------------
' просмотр и выбранном фолдере Outlook писем BTO и запись их в общий файл BTOmail.txt

Option Explicit

Const DirDBs = "C:\Users\Pavel_Khrapkin\Documents\Pavel\match\Match2.0\DBs\"

Sub BTO_Mail_trace()
'
' <*> BPO_Outlook_Mail_trace()  - Create file BTOmails.txt from selected Mails in Folder "BTO"
'
' Pavel Khrapkin 12.6.2012
'   13.11.12 - адаптация в Office-2010 и настройка по месту фолдера BTO

    Dim myNameSpace As Outlook.NameSpace
    Dim myItems As Outlook.Items
    Dim Mail As MailItem
    Dim fld As Folder
    Dim i As Integer
    
    
    Dim F
    
    Set myNameSpace = Application.GetNamespace("MAPI")
    Set F = myNameSpace.Folders
    
'---------------------------------------------------
'------   настраивается по месту фолдера BTO  ------
'   Set myItems = F.Item(2).Folders.Item(17).Items
' Означает "во втором PST в папке 17 все письма"
' Если папка с BTO расположена в другом месте, надо
' в отладчике найти ее место в строке ниже и подстроить.

    Set myItems = F.Item(3).Folders.Item(2).Folders.Item(2).Items
    myItems.Sort "ReceivedTime", True
    
    ChDir DirDBs
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
    ChDir DirDBs

    Open "BTOtest.txt" For Append As #1
    Print #1, "Subj text " & Timer
    Print #1, "Subj text " & Timer
    Print #1, "Subj text " & Timer
End Sub

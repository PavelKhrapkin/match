Attribute VB_Name = "BTO_Mail_tracking"
Option Explicit

Sub BTO_Mail_trace()
'
' <*> BPO_Mail_trace()
'
'    MsgBox "начало"
'
'    Dim FolderNm As Outlook.Folder
'    Dim Inbox As Folder
'
    Dim myNamespace As Outlook.NameSpace
    Dim Mail As MailItem
    Dim fld As Folder
    Dim i As Integer
    
    
    Dim F
    
    Set myNamespace = Application.GetNamespace("MAPI")
    Set F = myNamespace.Folders
    
'    MsgBox F.Item(2).Folders.Item(17)
    For Each Mail In F.Item(2).Folders.Item(17).Items
        BTO_MailHandle Mail.Subject, Mail.Body
    Next Mail
'    Set Application.ActiveExplorer.CurrentFolder = _
'            myNamespace.GetDefaultFolder(olFolderInbox)
'
'    For Each fld In Application.ActiveExplorer.CurrentFolder
'        MsgBox fld
'    Next fld
'
''    Application.ActiveExplorer.CurrentFolder.FolderPath
'
'    For Each Mail In Application.ActiveExplorer.CurrentFolder.Items
'        MsgBox Mail.Subject
'    Next Mail
    
End Sub
Sub BTO_MailHandle(ByVal MsgSubj, ByVal MsgBody)
'
' - BTO_MailHanle(Mail) - обработка письма Mail
'   10.6.12

    
    Dim IsSbs As Boolean
'    Dim Subj As String
'    MsgBox MsgSubj
    
'    Subj = Mail.Subject
    
    If InStr(MsgSubj, "Обновление по подписке") <> 0 Then
        MsgBox "<Sbs> CSD: подписка " & MsgSubj & vbCrLf & vbCrLf & MsgBody
    End If
End Sub
Sub CompanyChange()
    Dim ContactsFolder As Folder
    Set ContactsFolder = Session.GetDefaultFolder(olFolderContacts)
    MsgBox ("Contacts Found: " & ContactsFolder.Items.Count)
End Sub


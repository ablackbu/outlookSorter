Attribute VB_Name = "MyMailCommons"

' Generic function to locate a folder by the path
Function GetFolder(ByVal folderPath As String) As Outlook.folder
 Dim TestFolder As Outlook.folder
 Dim FoldersArray As Variant
 Dim i As Integer
 
 On Error GoTo GetFolder_Error
 If Left(folderPath, 2) = "\\" Then
 folderPath = Right(folderPath, Len(folderPath) - 2)
 End If
 'Convert folderpath to array
 FoldersArray = Split(folderPath, "\")
 Set TestFolder = Application.Session.folders.Item(FoldersArray(0))
 If Not TestFolder Is Nothing Then
 For i = 1 To UBound(FoldersArray, 1)
 Dim SubFolders As Outlook.folders
 Set SubFolders = TestFolder.folders
 Set TestFolder = SubFolders.Item(FoldersArray(i))
 If TestFolder Is Nothing Then
 Set GetFolder = Nothing
 End If
 Next
 End If
 'Return the TestFolder
 Set GetFolder = TestFolder
 Exit Function
 
GetFolder_Error:
 Set GetFolder = Nothing
 Exit Function
End Function

' Generic sorter that does all the sorting for the others
Function sortGenericVersion3(NumberOfDaysUntilSort As Integer, SenderName As String, folderPath As String)

    On Error Resume Next
    
    Debug.Print "DaysUntilSort: " & NumberOfDaysUntilSort & " SenderName: " & SenderName & " FolderPath: " & folderPath
    
    'TODO: Refactor me into the MyMailSorter.bas so we don't need config in multiple locations
    Set objInbox = MyMailCommons.GetFolder("\\email.user@domain.com\Inbox")
    Set destFolder = MyMailCommons.GetFolder(folderPath)
    
    Set colItems = objInbox.Items
    Dim senderSearchString As String
    senderSearchString = "[SenderName] = '" & SenderName & "'"
    Set objItem = colItems.Find(senderSearchString)
    
    Do While TypeName(objItem) <> "Nothing"
        Debug.Print "Message: " & objItem
        Dim days As Integer
        days = Abs(DateDiff("d", Now, objItem.ReceivedTime))
        Debug.Print "Days: " & days
        
        If days >= NumberOfDaysUntilSort Then
            objItem.Move destFolder
        End If
        
        'MsgBox "Sender: " & SenderName & " rec: " & objItem.ReceivedTime & " days: " & days
        
        Set objItem = colItems.FindNext
    Loop

End Function



' Generic marker that will mark everything in the path as read
Sub markFolderAsRead(folderPath As String)

    Set objFolder = MyMailCommons.GetFolder(folderPath)
    ' Debug.Print prints messages to the Immediate console accessed by ctrl+G
    For Each objMessage In objFolder.Items
        Debug.Print "Message: " & objMessage
        objMessage.UnRead = False
    Next
    
End Sub



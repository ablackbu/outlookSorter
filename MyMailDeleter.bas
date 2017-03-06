Attribute VB_Name = "MyMailDeleter"
Sub RemoveAllItemsAndFoldersInDeletedItems()
    Dim oDeletedItems As Outlook.folder
    Dim oFolders As Outlook.folders
    Dim oItems As Outlook.Items
    Dim i As Long
    'Obtain a reference to deleted items folder
    Set oDeletedItems = Application.Session.GetDefaultFolder(olFolderDeletedItems)
    Set oItems = oDeletedItems.Items
    For i = oItems.Count To 1 Step -1
        oItems.Item(i).Delete
    Next
    Set oFolders = oDeletedItems.folders
    For i = oFolders.Count To 1 Step -1
        oFolders.Item(i).Delete
    Next
End Sub



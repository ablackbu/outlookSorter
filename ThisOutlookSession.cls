VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Application_Startup()
    Call MyMailSorter.runAllSorts
    Call MyMailCommons.markFolderAsRead("\\email.user@domain.com\Deleted Items")
End Sub



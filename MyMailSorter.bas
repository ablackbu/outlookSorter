Attribute VB_Name = "MyMailSorter"
Sub runAllSorts()
   Dim DaysUntilSort As Integer
   DaysUntilSort = 5
   
   Dim MyEmail As String
   MyInbox = "\\email.user@domain.com\Inbox"
   
   ' A few different examples
   Call MyMailCommons.sortGenericVersion3(DaysUntilSort, "noreply@steampowered.com", MyInbox + "\Other\Steam")
   Call MyMailCommons.sortGenericVersion3(DaysUntilSort, "Amazon.com", MyInbox + "\Shopping\Amazon")

End Sub



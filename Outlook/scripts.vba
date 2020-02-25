
Public Sub SaveAsMHT()
On Error Resume Next

    Dim MyWindow As Outlook.Inspector
    Dim MyItem As MailItem
    Dim FilePath As String
    
    FilePath = "C:\Documentation\MailMemo\"
    Dim ItemName As String
    Set MyWindow = Application.ActiveInspector
    
    If TypeName(MyWindow) = "Nothing" Then
        'Get selection
        Set MyItem = Application.ActiveExplorer.Selection.Item(1)
    Else
        'Get active object
        Set MyItem = MyWindow.CurrentItem
    End If
        
    ItemName = MyItem.Subject
    ItemName = ReplaceIllegalCharacters(ItemName, " ") 'Remove illegal characters
    ItemName = Replace(ItemName, "  ", " ") 'Remove double spaces
    MyItem.SaveAs FilePath & ItemName & ".mht", olMHTML
 
    MsgBox ("Saved")
End Sub

Public Function ReplaceIllegalCharacters(strIn As String, strChar As String) As String
    Dim strSpecialChars As String
    Dim i As Long
    strSpecialChars = "~""#%&*:<>?{|}/\[]" & Chr(10) & Chr(13)

    For i = 1 To Len(strSpecialChars)
        strIn = Replace(strIn, Mid$(strSpecialChars, i, 1), strChar)
    Next

    ReplaceIllegalCharacters = strIn
End Function


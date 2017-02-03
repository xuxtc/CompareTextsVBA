VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim strWholeStory As String
Dim strActual As String, strExpected As String
Dim separator As Integer
Dim selectedRange As Range

Private Sub findDifference()

On Error Resume Next

getParagraph

For i = 1 To Len(strExpected)
    If Mid(strExpected, i, 1) <> Mid(strActual, i, 1) Then
        With ActiveDocument.Range(Start:=i, End:=i + 1)
            .Font.ColorIndex = wdRed
        End With
    End If
Next i

End Sub

Private Function getParagraph() As String
    Selection.WholeStory
    strWholeStory = removeSpecialCharacter(Trim(Selection.Text))
    
    separator = getSeparatorIndex(strWholeStory)
    
    strExpected = Mid(strWholeStory, 1, separator - 1)
    strActual = Mid(strWholeStory, separator + 3, Len(strWholeStory) - separator)
End Function

Private Function getSeparatorIndex(str_text As String) As Integer
    getSeparatorIndex = InStr(str_text, "&&&")
End Function

Private Function removeSpecialCharacter(str_in As String) As String
    str_in = Replace(str_in, Chr(1), "")
    removeSpecialCharacter = Trim(str_in)
End Function

Private Sub CommandButton21_Click()
    findDifference
End Sub

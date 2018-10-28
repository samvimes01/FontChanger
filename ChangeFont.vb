Sub ChangeEveryFifthWord()
Dim i As Long
Dim oWord As Range

Set oWord = ActiveDocument.Words.First
i = 1

Do While i <= ActiveDocument.Words.Count

While Trim(oWord.Text) Like "[ ]" Or oWord.Text = vbCr
If oWord.End = ActiveDocument.Words.Last.End Then Exit Sub
Set oWord = oWord.Next(wdWord)
Wend

If i Mod 2 = 0 Then
oWord.MoveEndWhile " "
oWord.Font.Name = "My Font IV"
Else
oWord.Font.Name = "My Font III"
End If
If i Mod 5 = 0 Then
oWord.InsertAfter " "
End If

If i Mod 8 = 0 Then
oWord.InsertAfter " "
End If

i = i + 1
Set oWord = oWord.Next(wdWord)
If oWord.End = ActiveDocument.Words.Last.End Then Exit Sub
DoEvents

Loop
End Sub
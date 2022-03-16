Attribute VB_Name = "Module1"
Sub RunTests()
Dim s As Shape
Set s = ActiveSheet.Shapes(Application.Caller)
Application.Run s.TextFrame.Characters.Text
End Sub

Private Sub test1()
    Dim out As New Collection
    out.Add uAnswer.Answer(, , "Choices choices")
    If IsObject(out(1)) Then
        MsgBox TypeName(out(1)) & vbTab & out(1).Address
    Else
        MsgBox TypeName(out(1)) & vbTab & out(1)
    End If
End Sub

Private Sub test2()
    Dim out As New Collection
    out.Add uAnswer.Answer(, Array("A", "B", 1, 2), "Choices choices")
    If IsObject(out(1)) Then
        MsgBox TypeName(out(1)) & vbTab & out(1).Address
    Else
        MsgBox TypeName(out(1)) & vbTab & out(1)
    End If
End Sub
Private Sub test3()
    Dim out As New Collection
    out.Add uAnswer.Answer(argTrueFalse + argYesNo)

    If IsObject(out(1)) Then
        MsgBox TypeName(out(1)) & vbTab & out(1).Address
    Else
        MsgBox TypeName(out(1)) & vbTab & out(1)
    End If
End Sub

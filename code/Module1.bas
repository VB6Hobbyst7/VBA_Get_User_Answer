Attribute VB_Name = "Module1"
Public Answer As Variant

Const ControlIDCheckBox = "Forms.CheckBox.1"
Const ControlIDComboBox = "Forms.ComboBox.1"
Const ControlIDCommandButton = "Forms.CommandButton.1"
Const ControlIDFrame = "Forms.Frame.1"
Const ControlIDImage = "Forms.Image.1"
Const ControlIDLabel = "Forms.Label.1"
Const ControlIDListBox = "Forms.ListBox.1"
Const ControlIDMultiPage = "Forms.MultiPage.1"
Const ControlIDOptionButton = "Forms.OptionButton.1"
Const ControlIDScrollBar = "Forms.ScrollBar.1"
Const ControlIDSpinButton = "Forms.SpinButton.1"
Const ControlIDTabStrip = "Forms.TabStrip.1"
Const ControlIDTextBox = "Forms.TextBox.1"
Const ControlIDToggleButton = "Forms.ToggleButton.1"

Sub test()
    Dim X As Variant
    GetAnswer Array("A", "B", 1, 2), "Choices choices"
    Dim out As Variant
    If IsObject(Answer) Then
        out = Answer.Address
    Else
        out = Answer
    End If
    MsgBox TypeName(Answer) & vbTab & out
End Sub

Sub GetAnswer(ExtraOptions As Variant, Optional Caption As String)
    Dim counter As Long
    counter = UBound(ExtraOptions) + 1
    Load uAnswer
    Dim i As Long
    Dim c As Control
    For i = 1 To counter
        Set c = uAnswer.Controls.Add(ControlIDOptionButton)
        c.Left = 6
        c.Top = 60 + (i * 18)
        c.Caption = ExtraOptions(i - 1)
        c.AutoSize = True
    Next
    If Len(Caption) > 0 Then uAnswer.Caption = Caption
    ResizeUserformToFitControls uAnswer
    uAnswer.Show
End Sub
Function InputBoxRange(Optional sTitle As String, Optional sPrompt As String) As Range
    On Error Resume Next
    Set InputBoxRange = Application.InputBox(Title:=sTitle, Prompt:=sPrompt, Type:=8, _
                                    Default:=IIf(TypeName(Selection) = "Range", Selection.Address, ""))
End Function
Sub ResizeUserformToFitControls(Form As Object)
    Dim ctr As MSForms.Control
    Dim myWidth
    myWidth = Form.InsideWidth
    For Each ctr In Form.Controls
        If ctr.Left + ctr.Width > myWidth Then myWidth = ctr.Left + ctr.Width
    Next
    Form.Width = myWidth + Form.Width - Form.InsideWidth        '+ 10
    Dim myHeight
    myHeight = Form.InsideHeight
    For Each ctr In Form.Controls
        If ctr.Top + ctr.Height > myHeight Then myHeight = ctr.Top + ctr.Height
    Next
    Form.Height = myHeight + Form.Height - Form.InsideHeight        '+ 10
End Sub

Function whichOption(Frame As Variant, controlType As String) As Variant
    Dim out As New Collection
    For Each Control In Frame.Controls
        If UCase(TypeName(Control)) = UCase(controlType) Then
            If Control.Value = True Then
                out.Add Control
            End If
        End If
    Next
    If out.Count = 1 Then
        Set whichOption = out(1)
    ElseIf out.Count > 1 Then
        Set whichOption = out
    End If
End Function



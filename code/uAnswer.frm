VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uAnswer 
   Caption         =   "UserForm1"
   ClientHeight    =   2760
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6960
   OleObjectBlob   =   "uAnswer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uAnswer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents Calendar1 As cCalendar
Attribute Calendar1.VB_VarHelpID = -1
Private ans As Variant

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

Enum AnswerType
    argInput = 2
    argYesNo = 4
    argTrueFalse = 8
    argDate = 16
    argRange = 32
End Enum

Private Sub UserForm_Activate()
ResizeUserformToFitControls Me
Me.Width = 195
If TextBox1.Visible = True Then TextBox1.SetFocus
End Sub

Private Sub UserForm_Initialize()

Set Calendar1 = New cCalendar
    With Calendar1
        .Add_Calendar_into_Frame Me.Frame1
        .UseDefaultBackColors = True
        .DayLength = 3
        .MonthLength = mlENShort
    End With

End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then cmdOK_Click
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Set Calendar1 = Nothing
End Sub

Public Function Answer(Optional AT As AnswerType = 999, Optional ExtraOptions As Variant, Optional Caption As String) As Variant

    If AT = 999 Then AT = argDate + argInput + argRange + argTrueFalse + argYesNo
    If AT And argDate Then oDate.Visible = True
    If AT And argInput Then oInput.Visible = True: TextBox1.Visible = True
    If AT And argRange Then oRange.Visible = True
    If AT And argTrueFalse Then oTrue.Visible = True: oFalse.Visible = True
    If AT And argYesNo Then oYes.Visible = True: oNo.Visible = True

    If IsArray(ExtraOptions) Then
        Dim counter As Long
        counter = UBound(ExtraOptions) + 1
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
    End If
    
    Me.Show
    If IsObject(ans) Then
        Set Answer = ans
    Else
        Answer = ans
    End If
    Unload Me
End Function

Private Sub cmdOK_Click()
ans = whichOption(Me, "OptionButton").Caption
If ans = "INPUT" Then
    ans = TextBox1.Text
ElseIf ans = "RANGE" Then
    Set ans = InputBoxRange
    GoTo NORMAL_EXIT
ElseIf ans = "vbYES" Then
    ans = vbYes
ElseIf ans = "vbNO" Then
    ans = vbNo
ElseIf ans = "TRUE" Then
    ans = True
ElseIf ans = "FALSE" Then
    ans = False
ElseIf ans = "DATE" Then
    ans = Calendar1.Value
End If

If TypeName(ans) = "Boolean" Then
        ans = CBool(ans)
ElseIf IsDate(ans) Then
    ans = CDate(ans)
ElseIf IsNumeric(ans) Then
        ans = CLng(ans)
ElseIf TypeName(ans) = "String" Then
    ans = CStr(ans)
End If

NORMAL_EXIT:
Me.Hide
End Sub

Private Sub cmdCancel_Click()
'Answer = vbCancel
Unload Me
End Sub

Private Function InputBoxRange(Optional sTitle As String, Optional sPrompt As String) As Range
    On Error Resume Next
    Set InputBoxRange = Application.InputBox(Title:=sTitle, Prompt:=sPrompt, Type:=8, _
                                    Default:=IIf(TypeName(Selection) = "Range", Selection.Address, ""))
End Function
Private Sub ResizeUserformToFitControls(Form As Object)
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

Private Function whichOption(Frame As Variant, controlType As String) As Variant
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


Private Sub cmdToday_Click()
    Calendar1.Year = Format(Date, "YYYY")
    Calendar1.Month = Format(Date, "MM")
    Calendar1.Day = Format(Date, "DD")
End Sub

Private Sub oDate_Change()
If oDate = True Then
cmdToday.Visible = True
Me.Width = 360
Else
Me.Width = 195
cmdToday.Visible = False
End If
End Sub


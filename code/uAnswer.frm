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


Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then cmdOK_Click
End Sub

Private Sub UserForm_Activate()
Me.Width = 195
Me.TextBox1.SetFocus
End Sub

Private Sub UserForm_Initialize()
Answer = ""

Set Calendar1 = New cCalendar
    With Calendar1
        .Add_Calendar_into_Frame Me.Frame1
        .UseDefaultBackColors = True
        .DayLength = 3
        .MonthLength = mlENShort
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Set Calendar1 = Nothing
End Sub

Private Sub cmdToday_Click()
    Calendar1.Year = Format(Date, "YYYY")
    Calendar1.Month = Format(Date, "MM")
    Calendar1.Day = Format(Date, "DD")
End Sub

Private Sub cmdOK_Click()

Answer = ""
Dim ans As Variant
ans = whichOption(Me, "OptionButton").Caption
If ans = "INPUT" Then
    ans = TextBox1.Text
ElseIf ans = "RANGE" Then
    Set Answer = InputBoxRange
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

Answer = ans

NORMAL_EXIT:
Unload Me
End Sub

Private Sub cmdCancel_Click()
Answer = vbCancel
Unload Me
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


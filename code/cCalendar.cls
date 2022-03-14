VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Put the following in a userform (needs a frame)

'////////////////////////////////
'Private WithEvents Calendar1 As cCalendar
'
'Private Sub Calendar1_DblClick()
'ActiveCell.Value = Calendar1.Value
'Unload Me
'End Sub
'
'Private Sub UserForm_Initialize()
'Set Calendar1 = New cCalendar
'With Calendar1
'    .Add_Calendar_into_Frame Me.Frame1
'    .UseDefaultBackColors = True
'    .DayLength = 3
'    .MonthLength = mlENShort
'End With
'End Sub
'
'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Set Calendar1 = Nothing
'End Sub

'///////////////////////////////////

Option Explicit

'#########################
'# Team authors:          #
'# Krisztina Szabó         #
'# Gábor Madács             #
'# Roberto Mensa (nick r)    #
'##############################

'# Event Triggered By Main Object
Public Event AfterUpdate()
Public Event BeforeUpdate(ByRef Cancel As Integer)
Public Event Click()
Public Event DblClick()

'# Members for Main Object
Private WithEvents CBxY As MSForms.ComboBox
Attribute CBxY.VB_VarHelpID = -1
Private WithEvents CBxM As MSForms.ComboBox
Attribute CBxM.VB_VarHelpID = -1

Private CLb As MSForms.Label
Private mDayButtons() As cCalendar
Private mLabelButtons() As cCalendar

Private PTitleNewFont As MSForms.NewFont
Private PDayNewFont As MSForms.NewFont
Private PGridNewFont As MSForms.NewFont
'# Members for Button Object
Private WithEvents CmB As MSForms.CommandButton
Attribute CmB.VB_VarHelpID = -1
Private CmBl As MSForms.Label
Private CmBlNum As MSForms.Label
Private mcMain As cCalendar

'# For Properties
Private lPFontSize As Long
Private lPMonthLength As calMonthLength
Private lPDayLength As Long
Private bPYearFirst As Boolean
Private lPTitleFontColor As Long
Private lPGridFontColor As Long
Private lPDayFontColor As Long
Private lPFirstDay As calDayOfWeek
Private dValue As Date
Private lPBackColor As Long
Private lPMonth As Long
Private lPYear As Long
Private lPDay As Long
Private lPHeaderBackColor As Long
Private lPUseDefaultBackColors  As Boolean
Private bPVisible As Boolean
Private sPHeight As Single
Private sPWidth As Single
Private sPTop As Single
Private sPLeft As Single
Private lPSaturdayBackColor As Long
Private lPSundayBackColor As Long
Private sPControlTipText As String
Private bPTabStop As Boolean
Private lPTabIndex As Long
Private sPTag As String

Private bPShowDays As Boolean
Private bPShowTitle As Boolean
Private bPShowDateSelectors As Boolean
Private bPValueIsNull As Boolean

Private Const cDayFontColorSelected As Long = &H80000012
Private Const cDayFontColorInactive As Long = &H80000011
Private Const cBackColorInactive As Long = &H80000011
Private Const cDefaultWidth As Single = 216
Private Const cDefaultHeight As Single = 144

Public Enum calDayOfWeek
    dwMonday = 1
    dwTuesday = 2
    dwWednesday = 3
    dwThursday = 4
    dwFriday = 5
    dwSaturday = 6
    dwSunday = 7
End Enum

Public Enum calMonthLength
    mlLocalLong = 0
    mlLocalShort = 1
    mlENLong = 2
    mlENShort = 3
End Enum




'###########################
'# Properties for Main object
Private Sub AAA__Properties_Blank_for_compatibility()
End Sub

Public Property Get GridCellEffect() As Long
'Property Blank - not work
'Determines the effect used to display the grid.
End Property

Public Property Get GridLinesColor() As Long
'Property Blank - not work
'Determines the color used to display the lines in the grid.
End Property

Public Property Get ShowHorizontalGrid() As Boolean
'Property Blank - not work
'Specifies whether the calendar display horizontal gridlines.
End Property

Public Property Get ShowVerticalGrid() As Boolean
'Property Blank - not work
'Specifies whether to display vertical gridlines.
End Property

Public Property Get HelpContextID() As Long
'Property Blank - not work
'Specifies whether to display vertical gridlines.
End Property


'###########################
'# Properties for Main object
Private Sub AAA__Properties_Book()
End Sub

Public Property Get Tag() As String
    Tag = sPTag
End Property

Public Property Let Tag(sTag As String)
    sPTag = sTag
End Property

Public Property Get Parent() As Control
    If bInit Then
        Set Parent = CBxY.Parent.Parent
    Else
        Set Parent = Nothing
    End If
End Property

Public Property Get ValueIsNull() As Boolean
    ValueIsNull = bPValueIsNull
End Property

Public Property Let ValueIsNull(ByVal bValueIsNull As Boolean)
    bPValueIsNull = bValueIsNull
    If bInit Then
        Value = Value
    End If
End Property

Public Property Get ShowTitle() As Boolean
    ShowTitle = bPShowTitle
End Property

Public Property Let ShowTitle(ByVal bShowTitle As Boolean)
    bPShowTitle = bShowTitle
    If bInit Then
        CLb.Visible = bPShowTitle
        Move
    End If
End Property

Public Property Get ShowDays() As Boolean
    ShowDays = bPShowDays
End Property

Public Property Let ShowDays(ByVal bShowDays As Boolean)
    Dim i As Long
    bPShowDays = bShowDays
    If bInit Then
        For i = 0 To 6
            mLabelButtons(i).Obj_CmBl.Visible = bShowDays
        Next
        Move
    End If
End Property

Public Property Get ShowDateSelectors() As Boolean
    ShowDateSelectors = bPShowDateSelectors
End Property

Public Property Let ShowDateSelectors(ByVal bShowDateSelectors As Boolean)
    bPShowDateSelectors = bShowDateSelectors
    If bInit Then
        CBxY.Visible = bShowDateSelectors
        CBxM.Visible = bShowDateSelectors
        Move
    End If
End Property

Public Property Get TabIndex() As Long
    TabIndex = lPTabIndex
End Property

Public Property Let TabIndex(ByVal lTabIndex As Long)
    lPTabIndex = lTabIndex
    If bInit Then
        CBxY.Parent.TabIndex = lTabIndex
    End If
End Property

Public Property Get TabStop() As Boolean
    TabStop = bPTabStop
End Property

Public Property Let TabStop(ByVal bTabStop As Boolean)
    bPTabStop = bTabStop
    If bInit Then
        CBxY.Parent.TabStop = bTabStop
    End If
End Property

Public Property Get ControlTipText() As String
    ControlTipText = sPControlTipText
End Property

Public Property Let ControlTipText(ByVal sControlTipText As String)
    Dim i As Long
    sPControlTipText = sControlTipText
    If bInit Then
        For i = 0 To 6
            mLabelButtons(i).Obj_CmBl.ControlTipText = sControlTipText
        Next
        For i = 0 To 41
            mDayButtons(i).Obj_Cmb.ControlTipText = sControlTipText
        Next
        CBxM.ControlTipText = sControlTipText
        CBxY.ControlTipText = sControlTipText
        CLb.ControlTipText = sControlTipText
        'CBxY.Parent.ControlTipText = sControlTipText
    End If
End Property

Public Property Get GridFont() As MSForms.NewFont
    Set GridFont = PGridNewFont
End Property

Public Property Set GridFont(ByRef clGridNewFont As MSForms.NewFont)
    Set PGridNewFont = clGridNewFont
End Property

Public Property Get DayFont() As MSForms.NewFont
    Set DayFont = PDayNewFont
End Property

Public Property Set DayFont(ByRef clDayNewFont As MSForms.NewFont)
    Set PDayNewFont = clDayNewFont
End Property

Public Property Get TitleFont() As MSForms.NewFont
    Set TitleFont = PTitleNewFont
End Property

Public Property Set TitleFont(ByRef clTitleNewFont As MSForms.NewFont)
    Set PTitleNewFont = clTitleNewFont
End Property

Public Property Get Visible() As Boolean
    Visible = bPVisible
End Property

Public Property Let Visible(ByVal bVisible As Boolean)
    bPVisible = bVisible
    If bInit Then
        CBxY.Parent.Visible = bVisible
    End If
End Property

Public Property Get Left() As Single
    Left = sPLeft
End Property

Public Property Let Left(ByVal sLeft As Single)
    sPLeft = sLeft
    If bInit Then
        CBxY.Parent.Left = sLeft
    End If
End Property

Public Property Get Top() As Single
    Top = sPTop
End Property

Public Property Let Top(ByVal ssTop As Single)
    sPTop = ssTop
    If bInit Then
        CBxY.Parent.Top = ssTop
    End If
End Property

Public Property Get Height() As Single
    Height = sPHeight
End Property

Public Property Let Height(ByVal sHeight As Single)
    sPHeight = sHeight
    If bInit Then
        CBxY.Parent.Height = sHeight
        Move
    End If
End Property


Public Property Get Width() As Single
    Width = sPWidth
End Property

Public Property Let Width(ByVal sWidth As Single)
    'sWidth = Zero_Negative_Value(sWidth)
    sPWidth = sWidth
    If bInit Then
        CBxY.Parent.Width = sWidth
        Move
    End If
End Property

Public Property Get BackColor() As Long
    BackColor = lPBackColor
End Property

Public Property Let BackColor(ByVal lBackColor As Long)
    lPBackColor = lBackColor
    If bInit Then
        CBxY.Parent.BackColor = lBackColor
    End If
End Property

Public Property Get HeaderBackColor() As Long
    HeaderBackColor = lPHeaderBackColor
End Property

Public Property Let HeaderBackColor(ByVal lHeaderBackColor As Long)
Dim i As Long
    lPHeaderBackColor = lHeaderBackColor
    UseDefaultBackColors = False
End Property

Public Property Get UseDefaultBackColors() As Boolean
    UseDefaultBackColors = lPUseDefaultBackColors
End Property

Public Property Let UseDefaultBackColors(ByVal lUseDefaultBackColors As Boolean)
    lPUseDefaultBackColors = lUseDefaultBackColors
    If bInit Then
        Refresh
    End If
End Property

Public Property Get SaturdayBackColor() As Long
    SaturdayBackColor = lPSaturdayBackColor
End Property

Public Property Let SaturdayBackColor(ByVal lSaturdayBackColor As Long)
    lPSaturdayBackColor = lSaturdayBackColor
    UseDefaultBackColors = False
End Property

Public Property Get SundayBackColor() As Long
    SundayBackColor = lPSundayBackColor
End Property

Public Property Let SundayBackColor(ByVal lSundayBackColor As Long)
    lPSundayBackColor = lSundayBackColor
    'If bInit Then
    '    Refresh
    'End If
    UseDefaultBackColors = False

End Property

Public Property Get FirstDay() As calDayOfWeek
    FirstDay = lPFirstDay
End Property

Public Property Let FirstDay(ByVal vbFirstDay As calDayOfWeek)
    Dim i As Long, v
    
    Select Case vbFirstDay
        Case 1 To 7
        Case Else
            vbFirstDay = 1
    End Select
    
    lPFirstDay = vbFirstDay
    If bInit Then
        v = fWeekdayName(CInt(lPDayLength))
        For i = 0 To 6
            mLabelButtons(i).Obj_CmBl.Caption = v(((i + vbFirstDay - 1) Mod 7))
        Next
        Refresh
    End If
End Property

Public Property Get DayFontColor() As Long
    DayFontColor = lPDayFontColor
End Property

Public Property Let DayFontColor(ByVal lFontColor As Long)
    Dim i As Long
    
    lPDayFontColor = lFontColor
    If bInit Then
        For i = 0 To 6
            mLabelButtons(i).Obj_CmBl.ForeColor = lFontColor
        Next
    End If
End Property

Public Property Get GridFontColor() As Long
    GridFontColor = lPGridFontColor
End Property

Public Property Let GridFontColor(ByVal lFontColor As Long)
    Dim i As Long
    lPGridFontColor = lFontColor
    If bInit Then
        Refresh
    End If
End Property

Public Property Let TitleFontColor(ByVal lFontColor As Long)
    lPTitleFontColor = lFontColor
    If bInit Then
        CLb.ForeColor = lFontColor
    End If
End Property

Public Property Get TitleFontColor() As Long
    TitleFontColor = lPTitleFontColor
End Property

Public Property Get Month() As Long
    Month = lPMonth
End Property

Public Property Let Month(ByVal lMonth As Long)
    If lMonth = 0 Then
        Value = Empty
    Else
        If lMonth < 0 Then lMonth = lPMonth
        lMonth = fMin(lMonth, 12)
        Value = SumMonthsToDate(dValue, lMonth - lPMonth)
    End If
    lPMonth = lMonth
End Property

Public Property Get Year() As Long
    Year = lPYear
End Property

Public Property Let Year(ByVal lYear As Long)
    If lYear = 0 Then
        Value = Empty
    Else
        Value = VBA.DateSerial(CheckYear(lYear), VBA.Month(dValue), VBA.Day(dValue))
    End If
    lPYear = lYear
End Property

Public Property Get Day() As Long
    Day = lPDay
End Property

Public Property Let Day(ByVal lDay As Long)
    If lDay = 0 Then
        Value = Empty
    Else
        If lDay < 0 Then lDay = lPDay
        lDay = fMin(lDay, VBA.Day(VBA.DateSerial(VBA.Year(dValue), VBA.Month(dValue) + 1, 0)))
        Value = VBA.DateSerial(VBA.Year(dValue), VBA.Month(dValue), lDay)
    End If
    lPDay = lDay
End Property

Public Property Get Value() As Variant
    If bPValueIsNull Then
        Value = Empty
    Else
        Value = dValue
    End If
End Property

Public Property Let Value(ByVal newDate As Variant)
    Dim Cancel As Integer '*** Integer for backward compatibility
    
    If CheckValue(newDate) = False Then newDate = Empty

    RaiseEvent BeforeUpdate(Cancel)
    
    If Cancel = 0 Then 'Not canceled.

        If bInit And Not IsEmpty(newDate) Then
            CBxY.ListIndex = VBA.Year(newDate) - 1904
            CBxM.ListIndex = VBA.Month(newDate) - 1
        End If
        
        If (bPValueIsNull = IsEmpty(newDate)) Or (newDate <> dValue) Then
            If Not IsEmpty(newDate) Then
                dValue = newDate
            End If
            bPValueIsNull = IsEmpty(newDate)
            
            If bInit Then
                Refresh
            End If
        End If
        
        RaiseEvent AfterUpdate
    End If
End Property

Public Property Get DayLength() As calMonthLength
    DayLength = lPDayLength
End Property

Public Property Let DayLength(ByVal bDayLength As calMonthLength)
    Dim i As Long, v

    lPDayLength = bDayLength
    If bInit Then
        v = fWeekdayName(bDayLength)
        For i = 0 To 6
            mLabelButtons(i).Obj_CmBl.Caption = v(((i + lPFirstDay - 1) Mod 7))
        Next
    End If
End Property

Public Property Get MonthLength() As calMonthLength
    MonthLength = lPMonthLength
End Property

Public Property Let MonthLength(ByVal iMonthLength As calMonthLength)
    Dim i As Long, m

    lPMonthLength = iMonthLength

    If bInit Then
        CBxM.List = fMonthName(CLng(iMonthLength))
        Value = Value
    End If
End Property

Public Property Get YearFirst() As Boolean
    YearFirst = bPYearFirst
End Property

Public Property Let YearFirst(ByVal bYearFirst As Boolean)
    bPYearFirst = bYearFirst
    RenderLabel
End Property


'###########################
'# Properties for Day button objects
Private Sub AAA__Properties_for_Day_button_Book()
End Sub

Public Property Set Main(ByVal theMain As cCalendar)
    Set mcMain = theMain
End Property

Private Property Get Main() As cCalendar
    Set Main = mcMain
End Property

Public Property Get Obj_Cmb() As MSForms.CommandButton
    Set Obj_Cmb = CmB
End Property

Public Property Set Obj_Cmb(ByVal vNewValue As MSForms.CommandButton)
    Set CmB = vNewValue
End Property

Public Property Get Obj_CmBl() As MSForms.Label
    Set Obj_CmBl = CmBl
End Property

Public Property Set Obj_CmBl(ByVal vNewValue As MSForms.Label)
    Set CmBl = vNewValue
End Property

Public Property Set Obj_CmBlNum(ByVal vNewValue As MSForms.Label)
    Set CmBlNum = vNewValue
End Property

Public Property Get Obj_CmBlNum() As MSForms.Label
    Set Obj_CmBlNum = CmBlNum
End Property


'###########################
'# Public Methods
Private Sub AAA_Methods_Book()
End Sub

Public Sub AboutBox()
    MsgBox "Autori: r, Kris, Gabor"
End Sub

Public Sub Add( _
    ByVal fForm As MSForms.UserForm)

    Dim cFrame As MSForms.Frame
    Set cFrame = fForm.Controls.Add("Forms.Frame.1")
    
    With cFrame
        .Width = IIf(sPWidth < 0, cDefaultWidth, sPWidth)
        .Height = IIf(sPHeight < 0, cDefaultHeight, sPHeight)
    End With
    
    Add_Calendar_into_Frame cFrame
    
End Sub

Public Sub Add_Calendar_into_Frame(ByVal cFrame As MSForms.Frame)
    Dim i As Long
    Dim v(199)
    Dim w
    Dim dTemp As Date
    
    For i = 0 To 199
        v(i) = CStr(1904 + i)
    Next
    
    With cFrame
        .BackColor = BackColor
        .Caption = ""
        .SpecialEffect = 0
        '.Top = IIf(sPTop = -1, .Top, sPTop)
        '.Left = IIf(sPLeft = -1, .Left, sPLeft)
        '.Width = IIf(sPWidth < 0, .Width, sPWidth)
        '.Height = IIf(sPHeight < 0, .Height, sPHeight)
        .Visible = bPVisible
        'Top = .Top
        'Left = .Left
        'Width = .Width
        'Height = .Height
    End With
    
    
    'Add this first, for proper taborder (Need TabStop.)
    Set CLb = cFrame.Controls.Add("Forms.Label.1")
    Set CBxY = cFrame.Controls.Add("Forms.ComboBox.1")
    Set CBxM = cFrame.Controls.Add("Forms.ComboBox.1")
    
    ReDim mLabelButtons(6)
    ReDim mDayButtons(41)
    w = fWeekdayName(CInt(lPDayLength))
    
    For i = 0 To 6
        Set mLabelButtons(i) = New cCalendar
        Set mLabelButtons(i).Main = Me
        Set mLabelButtons(i).Obj_CmBl = cFrame.Controls.Add("Forms.Label.1")
        With mLabelButtons(i).Obj_CmBl
            .Caption = w(((i + lPFirstDay - 1) Mod 7))
            .ForeColor = DayFontColor
            .TextAlign = fmTextAlignCenter
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = &H80000010 'Button shadow  &H80000015 'Button dark shadow
            '.SpecialEffect = fmSpecialEffectEtched
            If HeaderBackColor = -1 Then
                .BackColor = cBackColorInactive
                .BackStyle = fmBackStyleTransparent
            Else
                .BackColor = HeaderBackColor
                .BackStyle = fmBackStyleOpaque
            End If
        End With
    Next
            
    For i = 0 To 41
        Set mDayButtons(i) = New cCalendar
        Set mDayButtons(i).Main = Me
        
        Set mDayButtons(i).Obj_CmBl = cFrame.Controls.Add("Forms.Label.1")
        
        Set mDayButtons(i).Obj_CmBlNum = cFrame.Controls.Add("Forms.Label.1")
        With mDayButtons(i).Obj_CmBlNum
            .TextAlign = fmTextAlignCenter
            .BackStyle = fmBackStyleTransparent
        End With
        
        Set mDayButtons(i).Obj_Cmb = cFrame.Controls.Add("Forms.CommandButton.1")
        With mDayButtons(i).Obj_Cmb
            .BackStyle = fmBackStyleTransparent
        End With
    Next
    
    With CBxY
        .ListRows = 5
        .List = v
        .ListIndex = VBA.Year(dValue) - 1904
        .ShowDropButtonWhen = fmShowDropButtonWhenFocus
        .Font.Bold = True
        .MatchRequired = True
    End With

    With CBxM
        .ListRows = 12
        .List = fMonthName(lPMonthLength)
        .ListIndex = VBA.Month(dValue) - 1
        .ShowDropButtonWhen = fmShowDropButtonWhenFocus
        .Font.Bold = True
        .MatchRequired = True
    End With
    
    With CLb
        .ForeColor = TitleFontColor
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleTransparent
    End With
    
    ApplyFontChanges
    
    Refresh_Properities
    
    Move
    
End Sub

Private Sub ApplyFontChanges()
    Dim i As Long

    If Not PDayNewFont Is Nothing Then
        For i = 0 To 6
            With mLabelButtons(i).Obj_CmBl
                If .Font.Bold <> DayFont.Bold Then _
                    .Font.Bold = DayFont.Bold
                If .Font.Weight <> DayFont.Weight Then _
                    .Font.Weight = DayFont.Weight
                If .Font.Charset <> DayFont.Charset Then _
                    .Font.Charset = DayFont.Charset
                If .Font.Italic <> DayFont.Italic Then _
                    .Font.Italic = DayFont.Italic
                If .Font.Name <> DayFont.Name Then _
                    .Font.Name = DayFont.Name
                If .Font.size <> DayFont.size Then _
                    .Font.size = DayFont.size
                If DayFont.Strikethrough Then _
                    .Font.Strikethrough = True
                If DayFont.Underline Then _
                    .Font.Underline = True
            End With
        Next
    End If
            
    If Not PGridNewFont Is Nothing Then
        For i = 0 To 41
            With mDayButtons(i).Obj_CmBlNum
                If .Font.Bold <> GridFont.Bold Then _
                    .Font.Bold = GridFont.Bold
                If .Font.Weight <> GridFont.Weight Then _
                    .Font.Weight = GridFont.Weight
                If .Font.Charset <> GridFont.Charset Then _
                    .Font.Charset = GridFont.Charset
                If .Font.Italic <> GridFont.Italic Then _
                    .Font.Italic = GridFont.Italic
                If .Font.Name <> GridFont.Name Then _
                    .Font.Name = GridFont.Name
                If .Font.size <> GridFont.size Then _
                    .Font.size = GridFont.size
                If GridFont.Strikethrough Then _
                    .Font.Strikethrough = True
                If GridFont.Underline Then _
                    .Font.Underline = True
            End With
        Next
    End If
    
    If Not PTitleNewFont Is Nothing Then
        With CLb
                If .Font.Bold <> TitleFont.Bold Then _
                    .Font.Bold = TitleFont.Bold
                If .Font.Weight <> TitleFont.Weight Then _
                    .Font.Weight = TitleFont.Weight
                If .Font.Charset <> TitleFont.Charset Then _
                    .Font.Charset = TitleFont.Charset
                If .Font.Italic <> TitleFont.Italic Then _
                    .Font.Italic = TitleFont.Italic
                If .Font.Name <> TitleFont.Name Then _
                    .Font.Name = TitleFont.Name
                If .Font.size <> TitleFont.size Then _
                    .Font.size = TitleFont.size
                If TitleFont.Strikethrough Then _
                    .Font.Strikethrough = True
                If TitleFont.Underline Then _
                    .Font.Underline = True
        End With
    End If

End Sub

Public Sub Move( _
        Optional vLeft, _
        Optional vTop, _
        Optional vWidth, _
        Optional vHeight, _
        Optional vLayout)
        
    Dim i As Long, l As Currency, b As Currency, lc As Currency, bc As Currency
    Dim t As Long, b_ym As Currency, b_combo_m As Currency
    
    Const h_combo As Long = 16
    Const b_combo_y As Long = 42
    b_combo_m = IIf(lPMonthLength = mlENShort Or lPMonthLength = mlLocalShort, 42, 66) '66
    b_ym = b_combo_y + 2 + b_combo_m
    
    If bInit Then
        t = IIf(ShowDays, 7, 6)
        
        With CBxY.Parent 'Frame
            sPTop = IIf(IsMissing(vTop), IIf(Top = -1, .Top, Top), vTop)
            sPLeft = IIf(IsMissing(vLeft), IIf(Left = -1, .Left, Left), vLeft)
            sPHeight = IIf(IsMissing(vHeight), IIf(Height = -1, .Height, Height), vHeight)
            sPWidth = IIf(IsMissing(vWidth), IIf(Width = -1, .Width, Width), vWidth)
            
            l = Height
            b = Width
            l = Zero_Negative_Value(l - IIf(ShowTitle Or ShowDateSelectors, h_combo, 0) - 1)
            lc = CCur(l / t)
            bc = CCur(b / 7)
            b = bc * 7
            Debug.Print bc
            Debug.Print lc
            Debug.Print b
            Debug.Print l
            
        End With
        
        If ShowTitle Then
            With CLb
                .Width = Zero_Negative_Value(IIf(ShowDateSelectors, b - b_ym, b))
                .Height = h_combo
                .Left = 0
            End With
        End If
        
        If ShowDateSelectors Then
            With CBxY
                .Width = b_combo_y
                .Height = h_combo
                .Left = IIf(ShowTitle, CLb.Width, Int((b - b_ym) / 2)) + _
                       IIf(YearFirst, 0, b_combo_m + 2)
            End With
        
            With CBxM
                .Width = b_combo_m
                .Height = h_combo
                .Left = IIf(ShowTitle, CLb.Width, Int((b - b_ym) / 2)) + _
                       IIf(YearFirst, b_combo_y + 2, 0)
            End With
        End If
        If ShowDays Then
            For i = 0 To 6
                With mLabelButtons(i).Obj_CmBl
                    .Top = IIf(ShowTitle Or ShowDateSelectors, h_combo + 2, 0)
                    .Left = (i Mod 7) * bc - IIf(i > 0, 1, 0)
                    .Height = lc
                    .Width = bc + IIf(i > 0, 1, 0)
                End With
            Next
        End If
        For i = 0 To 41
            With mDayButtons(i).Obj_Cmb
                .Top = Int(i / 7) * lc + _
                       IIf(ShowTitle Or ShowDateSelectors, h_combo + 2, 0) + _
                       IIf(ShowDays, lc, 0)
                .Left = (i Mod 7) * bc
                .Height = lc
                .Width = bc
            End With
            With mDayButtons(i).Obj_CmBl
                .Top = mDayButtons(i).Obj_Cmb.Top
                .Left = mDayButtons(i).Obj_Cmb.Left
                .Height = mDayButtons(i).Obj_Cmb.Height
                .Width = mDayButtons(i).Obj_Cmb.Width
            End With
            
            With mDayButtons(i).Obj_CmBlNum
                .Top = Int(i / 7) * lc + _
                       IIf(ShowTitle Or ShowDateSelectors, h_combo, 0) + _
                       IIf(ShowDays, lc, 0) + 6
                .Left = (i Mod 7) * bc + 3
                .Height = Zero_Negative_Value(lc - 6)
                .Width = Zero_Negative_Value(bc - 6)
            End With

        Next
        
    Else
        sPHeight = IIf(IsMissing(Height), cDefaultHeight, Height)
        sPWidth = IIf(IsMissing(Width), cDefaultWidth, Width)
    End If
End Sub

Public Sub NextDay()
    Dim d As Date
    d = dValue + 1
    d = VBA.DateSerial(CheckYear(VBA.Year(d)), VBA.Month(d), VBA.Day(d))
    Value = d
End Sub

Public Sub NextWeek()
    Dim d As Date
    d = dValue + 7
    d = VBA.DateSerial(CheckYear(VBA.Year(d)), VBA.Month(d), VBA.Day(d))
    Value = d
End Sub

Public Sub NextMonth()
    Value = SumMonthsToDate(dValue, 1)
End Sub

Public Sub NextYear()
    Dim d As Date
    d = VBA.DateSerial(CheckYear(VBA.Year(dValue) + 1), VBA.Month(dValue), VBA.Day(dValue))
    Value = d
End Sub

Public Sub PreviousDay()
    Dim d As Date
    d = dValue - 1
    d = VBA.DateSerial(CheckYear(VBA.Year(d)), VBA.Month(d), VBA.Day(d))
    Value = d
End Sub

Public Sub PreviousWeek()
    Dim d As Date
    d = dValue - 7
    d = VBA.DateSerial(CheckYear(VBA.Year(d)), VBA.Month(d), VBA.Day(d))
    Value = d
End Sub

Public Sub PreviousMonth()
    Value = SumMonthsToDate(dValue, -1)
End Sub

Public Sub PreviousYear()
    Dim d As Date
    d = VBA.DateSerial(CheckYear(VBA.Year(dValue) - 1), VBA.Month(dValue), VBA.Day(dValue))
    Value = d
End Sub

Public Sub Today()
    Value = VBA.Date
End Sub

Public Sub Refresh()
    Refresh_Panel VBA.Month(dValue), VBA.Year(dValue)
    ApplyFontChanges
End Sub


'###########################
Private Sub AAA_Event_Book()
End Sub
'# Events for Main Object Components
'###########################

Private Sub CBxY_Change()
    RenderLabel
    Refresh_Panel CBxM.ListIndex + 1, CBxY.ListIndex + 1904
End Sub

Private Sub CBxM_Change()
    RenderLabel
    Refresh_Panel CBxM.ListIndex + 1, CBxY.ListIndex + 1904
End Sub

Private Sub CmB_Click()
    Main.Value = dValue
    Main.Event_click = True
End Sub

Private Sub CmB_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Main.Event_DblClick = True
End Sub


Private Sub CmB_KeyDown( _
    ByVal KeyCode As MSForms.ReturnInteger, _
    ByVal Shift As Integer)
    
    Dim newDate As Date

    '38 Up
    '37 Left
    '39 Right
    '40 Down
    
    newDate = dValue
    
    Select Case KeyCode
    Case 37
        newDate = newDate - 1
    Case 39
        newDate = newDate + 1
    Case 38
        newDate = newDate - 7
    Case 40
        newDate = newDate + 7
    Case 9
    End Select
    
    If newDate <> dValue Then
        Main.Value = newDate
        KeyCode = 0
    End If
End Sub


Private Sub Class_Initialize()
    bPShowDays = True
    bPShowTitle = True
    bPShowDateSelectors = True
    dValue = VBA.Date
    lPMonth = VBA.Month(VBA.Date)
    lPYear = VBA.Year(VBA.Date)
    lPDay = VBA.Day(VBA.Date)
    lPFontSize = 8
    lPMonthLength = 1
    lPDayLength = 1
    bPYearFirst = False
    lPTitleFontColor = &HA00000
    lPGridFontColor = &HA00000
    lPDayFontColor = &H0&
    lPFirstDay = 1
    lPBackColor = &H8000000F
    lPHeaderBackColor = 10053171 '&HFFAA99
    lPUseDefaultBackColors = True
    lPSaturdayBackColor = &H80000002
    lPSundayBackColor = &HFFAA99 '&H80000002
    bPVisible = True
    sPHeight = -1
    sPWidth = -1
    sPTop = -1
    sPLeft = -1
    sPControlTipText = ""
    
    Set TitleFont = New MSForms.NewFont
    With TitleFont
        .Name = "Arial"
        .size = lPFontSize + 4
        .Bold = True
    End With
    
    Set DayFont = New MSForms.NewFont
    With DayFont
        .Name = "Arial"
        .size = lPFontSize + 2
        .Bold = True
    End With
    
    Set GridFont = New MSForms.NewFont
    With GridFont
        .Name = "Arial"
        .size = lPFontSize
    End With

End Sub

Private Sub Class_Terminate()
    Erase mDayButtons
    Erase mLabelButtons
    Set mcMain = Nothing
    Set PTitleNewFont = Nothing
    Set PDayNewFont = Nothing
    Set PGridNewFont = Nothing
    Set CBxY = Nothing
    Set CBxM = Nothing
    Set CmB = Nothing
    Set CLb = Nothing
    Set CmBl = Nothing
End Sub

'###########################
'# Private Function
Private Sub AAA_Private_Function_Book()
End Sub

Private Function ArraY_Days(ByVal lMonth As Long, ByVal lYear As Long)
    Dim v(41) As Date, i As Long, g As Long, l As Long, p As Long
    
    i = VBA.DateTime.Weekday( _
        VBA.DateSerial(lYear, lMonth, 1), 1 + lPFirstDay Mod 7) - 1
    
    
    If i = 0 Then i = 7
    
    g = VBA.Day(VBA.DateSerial(lYear, lMonth + 1, 0)) + i
    
    p = 1
    For l = i To 0 Step -1
        v(l) = VBA.DateSerial(lYear, lMonth, p)
        p = p - 1
    Next
    
    p = 0
    For l = i To g
        p = p + 1
        v(l) = VBA.DateSerial(lYear, lMonth, p)
    Next
    
    For l = g To 41
        v(l) = VBA.DateSerial(lYear, lMonth, p)
        p = p + 1
    Next
    ArraY_Days = v
End Function

Private Sub RenderLabel()
    Dim b As Currency, b_ym As Currency, b_combo_m As Long
    
    Const b_combo_y As Long = 42
    b_combo_m = IIf(lPMonthLength = mlENShort Or lPMonthLength = mlLocalShort, 42, 66) '66
    b_ym = b_combo_y + 2 + b_combo_m
    
    If bInit Then
        b = CBxY.Parent.Width
        If bPYearFirst Then
            CLb.Caption = CBxY.Value & " " & CBxM.Value
        Else
            CLb.Caption = CBxM.Value & " " & CBxY.Value
        End If
        CLb.Width = Zero_Negative_Value(IIf(ShowDateSelectors, b - b_ym, b))
        CBxM.Width = b_combo_m
        CBxY.Left = IIf(ShowTitle, CLb.Width, CCur((b - b_ym) / 2)) + _
                       IIf(YearFirst, 0, b_combo_m + 2)
        CBxM.Left = IIf(ShowTitle, CLb.Width, CCur((b - b_ym) / 2)) + _
                       IIf(YearFirst, b_combo_y + 2, 0)
        'CBxY.Left = IIf(ShowTitle, CLb.Width, IIf(CLb.Width, Int(CLb.Width / 2), 0)) + _
        '           IIf(YearFirst, 0, b_combo_m + 2)
        '
        'CBxM.Left = IIf(ShowTitle, CLb.Width, IIf(CLb.Width, Int(CLb.Width / 2), 0)) + _
        '           IIf(YearFirst, b_combo_y + 2, 0)
    End If
End Sub

Private Function bInit() As Boolean
    If Not CBxY Is Nothing Then bInit = True
End Function


Private Function SumMonthsToDate(dDate As Date, Optional lMonth As Long = 1) As Date
    Dim d As Date
    
    d = VBA.DateSerial( _
            VBA.Year(dDate), _
            VBA.Month(dDate) + lMonth, _
            fMin( _
                VBA.Day(dDate), _
                VBA.Day( _
                    VBA.DateSerial( _
                    VBA.Year(dDate), _
                    VBA.Month(dDate) + 1 + VBA.Abs(lMonth), _
                    0))))
                    
    If d = VBA.DateSerial(CheckYear(VBA.Year(d)), VBA.Month(d), VBA.Day(d)) Then
        SumMonthsToDate = d
    Else
        SumMonthsToDate = dDate
    End If
End Function

Private Function fMin(vFirstValue, ParamArray vValues())
    Dim i As Long
    fMin = vFirstValue
    
    If IsMissing(vValues) = False Then
    For i = 0 To UBound(vValues)
        If fMin > vValues(i) Then
            fMin = vValues(i)
        End If
    Next
    End If
End Function

Private Function fMonthName(lIndex As Long)
    Dim m(11), i As Long, v
    lIndex = lIndex Mod 4
    If Int(lIndex / 2) Then
        If lIndex Mod 2 Then
            v = Array("Jan", "Feb", "Mar", "Apr", "May", _
                "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
        Else
            v = Array("January", "February", "March", _
                "April", "May", "June", "July", "August", _
                "September", "October", "November", "December")
        End If
        fMonthName = v
    Else
        For i = 0 To 11
            m(i) = VBA.Strings.MonthName(i + 1, lIndex Mod 2)
        Next
        fMonthName = m
    End If
End Function


Private Function fWeekdayName(lIndex As Long)
    Dim m(6), i As Long, v
    lIndex = lIndex Mod 4
    If Int(lIndex / 2) Then
        If lIndex Mod 2 Then
            v = Array("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
        Else
            v = Array("Monday", "Tuesday", "Wednestay", _
                "Thursday", "Friday", "Saturday", "Sunday")
        End If
        fWeekdayName = v
    Else
        For i = 0 To 6
            m(i) = VBA.Strings.WeekdayName(i + 1, lIndex Mod 2, vbMonday)
        Next
        fWeekdayName = m
    End If
End Function


Private Function CheckYear(ByVal lYear As Long) As Long
    Select Case lYear
    Case Is < 1904
        CheckYear = 1904
    Case 1904 To 2103
        CheckYear = lYear
    Case Else
        CheckYear = 2103
    End Select
End Function

'###########################
'# Private Sub
Private Sub AAA_Private_Sub_Book()
End Sub

Public Property Let Event_DblClick(ByVal X As Boolean)
    RaiseEvent DblClick
End Property

Private Property Get Event_DblClick() As Boolean
    Event_DblClick = False
End Property


Public Property Let Event_click(ByVal X As Boolean)
    RaiseEvent Click
End Property

Private Property Get Event_click() As Boolean
    Event_click = False
End Property

Private Sub Refresh_Properities()
    With Me
        .BackColor = .BackColor
        .ControlTipText = .ControlTipText
        .DayFontColor = .DayFontColor
        .DayLength = .DayLength
        .GridFontColor = .GridFontColor
        .MonthLength = .MonthLength
        If .UseDefaultBackColors = False Then
            .SaturdayBackColor = .SaturdayBackColor
            .SundayBackColor = .SundayBackColor
            .HeaderBackColor = .HeaderBackColor
        End If
        .ShowDateSelectors = .ShowDateSelectors
        .ShowDays = .ShowDays
        .ShowTitle = .ShowTitle
        .TabIndex = .TabIndex
        .TabStop = .TabStop
        .TitleFontColor = .TitleFontColor
        .ValueIsNull = .ValueIsNull
        .YearFirst = .YearFirst
    End With
End Sub

Private Sub Refresh_Selected_Day(ByVal dValue As Date)
    Dim i As Long, c As MSForms.Label
    For i = 0 To 41
        If mDayButtons(i).Value = dValue And Not bPValueIsNull Then
            On Error Resume Next
            mDayButtons(i).Obj_Cmb.SetFocus
            On Error GoTo 0
            With mDayButtons(i).Obj_CmBl
                .BackStyle = fmBackStyleOpaque
                .BackColor = cBackColorInactive
                .ForeColor = cDayFontColorSelected
            End With
            lPMonth = VBA.Month(dValue)
            lPYear = VBA.Year(dValue)
            lPDay = VBA.Day(dValue)
        End If
    Next

End Sub

Private Sub Refresh_Panel(ByVal lMonth As Long, ByVal lYear As Long)
    Dim v, i As Long, l As Long
    Dim iDay As Long
    
    If bInit Then
        v = ArraY_Days(lMonth, lYear)
        For i = 0 To 41
            mDayButtons(i).Value = v(i)
            With mDayButtons(i).Obj_CmBlNum
                If .Caption <> VBA.Day(v(i)) Then
                    .Caption = VBA.Day(v(i))
                End If
                If lMonth = VBA.Month(v(i)) Then
                    If .ForeColor <> GridFontColor Then
                        .ForeColor = GridFontColor
                    End If
                Else
                    If .ForeColor <> cDayFontColorInactive Then
                        .ForeColor = cDayFontColorInactive
                    End If
                End If
            End With
            With mDayButtons(i).Obj_CmBl
                If .BackStyle = fmBackStyleOpaque Then
                    .BackStyle = fmBackStyleTransparent
                End If
                If UseDefaultBackColors = False Then
                    iDay = VBA.DateTime.Weekday(v(i))
                    If iDay = vbSaturday Then
                        If .BackColor <> lPSaturdayBackColor Then
                            .BackColor = lPSaturdayBackColor
                        End If
                        If .BackStyle <> fmBackStyleOpaque Then
                            .BackStyle = fmBackStyleOpaque
                        End If
                    ElseIf iDay = vbSunday Then
                        If .BackColor <> lPSundayBackColor Then
                            .BackColor = lPSundayBackColor
                        End If
                        If .BackStyle <> fmBackStyleOpaque Then
                            .BackStyle = fmBackStyleOpaque
                        End If
                    End If
                End If
            End With
            If CheckValue(v(i)) = False Then
                mDayButtons(i).Obj_Cmb.Locked = True
            Else
                If mDayButtons(i).Obj_Cmb.Locked = True Then
                    mDayButtons(i).Obj_Cmb.Locked = False
                End If
            End If
        Next
        
        If UseDefaultBackColors = False Then
            For l = 0 To 6
                If mLabelButtons(l).Obj_CmBl.BackStyle = fmBackStyleTransparent Then _
                    mLabelButtons(l).Obj_CmBl.BackStyle = fmBackStyleOpaque
                If mLabelButtons(l).Obj_CmBl.BackColor <> lPHeaderBackColor Then _
                    mLabelButtons(l).Obj_CmBl.BackColor = lPHeaderBackColor
            Next
        Else
            For l = 0 To 6
                If mLabelButtons(l).Obj_CmBl.BackStyle = fmBackStyleOpaque Then _
                    mLabelButtons(l).Obj_CmBl.BackStyle = fmBackStyleTransparent
            Next
        End If
        
        If lMonth = VBA.Month(dValue) And lYear = VBA.Year(dValue) Then
            Refresh_Selected_Day dValue
        Else
            lPMonth = 0
            lPYear = 0
            lPDay = 0
        End If
    End If
End Sub

Private Function CheckValue(d) As Boolean
If VarType(d) = vbDate Then
    Select Case d
        Case 1462 To 74510
            CheckValue = CLng(d) = d
    End Select
End If
End Function

Private Function Zero_Negative_Value(sNumber As Single) As Single
If sNumber > 0 Then
    Zero_Negative_Value = sNumber
End If
End Function




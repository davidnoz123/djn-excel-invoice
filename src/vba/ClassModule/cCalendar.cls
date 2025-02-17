VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'###############################################################
'# Calendar Control Class v2.0.0                               #
'#                                                             #
'# Team authors:                                               #
'# Krisztina Szabó                                             #
'# Gábor Madács                                                #
'# Roberto Mensa (nick r)                                      #
'#                                                             #
'# https://sites.google.com/site/e90e50/calendar-control-class #
'#                                                             #
'#   The FrankensTeam                                          #
'###############################################################

'# Event Triggered By Main Object
Public Event AfterUpdate()
Public Event BeforeUpdate(ByRef Cancel As Integer)
Public Event Click()
Public Event DblClick()
Public Event KeyDown( _
    ByVal KeyCode As MSForms.ReturnInteger, _
    ByVal Shift As Integer)

'# Members for Main Object
Private WithEvents CBxY As MSForms.ComboBox
Attribute CBxY.VB_VarHelpID = -1
Private WithEvents CBxM As MSForms.ComboBox
Attribute CBxM.VB_VarHelpID = -1

Private CLb As MSForms.label
Private mDayButtons() As cCalendar
Private mLabelButtons() As cCalendar

Private PTitleNewFont As MSForms.NewFont
Private PDayNewFont As MSForms.NewFont
Private PGridNewFont As MSForms.NewFont
'# Members for Button Object
Private WithEvents CmB As MSForms.CommandButton
Attribute CmB.VB_VarHelpID = -1
Private CmBl As MSForms.label
Private CmBlNum As MSForms.label
Private mcMain As cCalendar

'# For Properties
Private lPFontSize As Long
Private lPMonthLength As calMonthLength
Private lPDayLength As Long
Private bPYearFirst As Boolean
Private lPTitleFontColor As OLE_COLOR
Private lPGridFontColor As OLE_COLOR
Private lPDayFontColor As OLE_COLOR
Private lPFirstDay As calDayOfWeek
Private dValue As Date
Private lPBackColor As OLE_COLOR
Private lPMonth As Long
Private lPYear As Long
Private lPDay As Long
Private lPHeaderBackColor As OLE_COLOR
Private lPUseDefaultBackColors  As Boolean
Private bPVisible As Boolean
Private sPHeight As Single
Private sPWidth As Single
Private sPTop As Single
Private sPLeft As Single
Private lPSaturdayBackColor As OLE_COLOR
Private lPSundayBackColor As OLE_COLOR
Private lPSelectedBackColor As OLE_COLOR
Private lPUnSelectableBackColor As OLE_COLOR
Private sPControlTipText As String
Private bPTabStop As Boolean
Private lPTabIndex As Long
Private sPTag As String

Private bPShowDays As Boolean
Private bPShowTitle As Boolean
Private bPShowDateSelectors As Boolean
Private bPValueIsNull As Boolean
Private bPRightToLeft As Boolean

Private bPMACFix As Boolean 'Fix MAC transparency errors
Private bPWeekdaySelectable As Boolean
Private bPSaturdaySelectable As Boolean
Private bPSundaySelectable As Boolean
Private bPValueSelectable As Boolean '(Used in buttons too).

Private maColoredArrayTable As Variant

Private Const cDayFontColorSelected As Long = &H80000012 'Button text - Black
Private Const cDayFontColorInactive As Long = &H80000011 'Disabled text - Dark gray
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

Public Enum calMonthLength '(Used for month and day names too.)
    mlLocalLong = 0 'Local name, long form
    mlLocalShort = 1 'Local name, short form
    mlENLong = 2 'English name, long form
    mlENShort = 3 'English name, short form
End Enum

Private Enum ColorCols 'ColoredDateArray fields
    ccColor = 1
    ccFormat = 2
    ccDateList = 3
    ccSelectable = 4
End Enum

'################################################
'# Properties for Main object - Not available

Public Property Get GridCellEffect() As Long
'Property Blank - not work
'Determines the effect used to display the grid.
End Property

Public Property Get GridLinesColor() As OLE_COLOR
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
'Specifies Help identifier
End Property


'###########################
'# Properties for Main object

Public Property Get Tag() As String
    Tag = sPTag
End Property

Public Property Let Tag(sTag As String)
    sPTag = sTag
End Property

Public Property Get Parent() As control
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
        Call Move
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
        Call Move
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
        Call Move
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
        Call Move
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
        Call Move
    End If
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = lPBackColor
End Property

Public Property Let BackColor(ByVal lBackColor As OLE_COLOR)
    lPBackColor = lBackColor
    If bInit Then
        CBxY.Parent.BackColor = lBackColor
    End If
End Property

Public Property Get HeaderBackColor() As OLE_COLOR
    HeaderBackColor = lPHeaderBackColor
End Property

Public Property Let HeaderBackColor(ByVal lHeaderBackColor As OLE_COLOR)
    lPHeaderBackColor = lHeaderBackColor
    UseDefaultBackColors = False
End Property

Public Property Get UseDefaultBackColors() As Boolean
    UseDefaultBackColors = lPUseDefaultBackColors
End Property

Public Property Let UseDefaultBackColors(ByVal lUseDefaultBackColors As Boolean)
    lPUseDefaultBackColors = lUseDefaultBackColors
    Call Refresh
End Property

Public Property Get SaturdayBackColor() As OLE_COLOR
    SaturdayBackColor = lPSaturdayBackColor
End Property

Public Property Let SaturdayBackColor(ByVal lSaturdayBackColor As OLE_COLOR)
    lPSaturdayBackColor = lSaturdayBackColor
    UseDefaultBackColors = False
End Property

Public Property Get SundayBackColor() As OLE_COLOR
    SundayBackColor = lPSundayBackColor
End Property

Public Property Let SundayBackColor(ByVal lSundayBackColor As OLE_COLOR)
    lPSundayBackColor = lSundayBackColor
    UseDefaultBackColors = False
End Property

Public Property Get SelectedBackColor() As OLE_COLOR
    SelectedBackColor = lPSelectedBackColor
End Property

Public Property Let SelectedBackColor(ByVal lSelectedBackColor As OLE_COLOR)
    lPSelectedBackColor = lSelectedBackColor
    Call Refresh
End Property

Public Property Get UnSelectableBackColor() As OLE_COLOR
    UnSelectableBackColor = lPUnSelectableBackColor
End Property

Public Property Let UnSelectableBackColor(ByVal lUnSelectableBackColor As OLE_COLOR)
    lPUnSelectableBackColor = lUnSelectableBackColor
    Call Refresh
End Property

Public Property Get SaturdaySelectable() As Boolean
    SaturdaySelectable = bPSaturdaySelectable
End Property

Public Property Let SaturdaySelectable(ByVal bSaturdaySelectable As Boolean)
    bPSaturdaySelectable = bSaturdaySelectable
    Call Refresh
End Property

Public Property Get SundaySelectable() As Boolean
    SundaySelectable = bPSundaySelectable
End Property

Public Property Let SundaySelectable(ByVal bSundaySelectable As Boolean)
    bPSundaySelectable = bSundaySelectable
    Call Refresh
End Property

Public Property Get WeekdaySelectable() As Boolean
    WeekdaySelectable = bPWeekdaySelectable
End Property

Public Property Let WeekdaySelectable(ByVal bWeekdaySelectable As Boolean)
    bPWeekdaySelectable = bWeekdaySelectable
    Call Refresh
End Property

Public Property Get FirstDay() As calDayOfWeek
    FirstDay = lPFirstDay
End Property

Public Property Let FirstDay(ByVal vbFirstDay As calDayOfWeek)
    Select Case vbFirstDay
        Case 1 To 7
        Case Else
            vbFirstDay = 1
    End Select
    
    lPFirstDay = vbFirstDay
    If bInit Then
        Call ApplyWeekDayLabelChanges
        Call Refresh
    End If
End Property

Public Property Get DayFontColor() As OLE_COLOR
    DayFontColor = lPDayFontColor
End Property

Public Property Let DayFontColor(ByVal lFontColor As OLE_COLOR)
    Dim i As Long
    
    lPDayFontColor = lFontColor
    If bInit Then
        For i = 0 To 6
            mLabelButtons(i).Obj_CmBl.ForeColor = lFontColor
        Next
    End If
End Property

Public Property Get GridFontColor() As OLE_COLOR
    GridFontColor = lPGridFontColor
End Property

Public Property Let GridFontColor(ByVal lFontColor As OLE_COLOR)
    lPGridFontColor = lFontColor
    Call Refresh
End Property

Public Property Let TitleFontColor(ByVal lFontColor As OLE_COLOR)
    lPTitleFontColor = lFontColor
    If bInit Then
        CLb.ForeColor = lFontColor
    End If
End Property

Public Property Get TitleFontColor() As OLE_COLOR
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
    If bPValueIsNull Or Not bPValueSelectable Then
        Value = Empty
    Else
        Value = dValue
    End If
End Property

Public Property Let Value(ByVal newDate As Variant)
    Dim Cancel As Integer '*** Integer for backward compatibility
    
    If CheckValue(newDate) = False Then newDate = Empty

    RaiseEvent BeforeUpdate(Cancel) '(Even if unselectable - for navigation.)
    
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
            
            Call Refresh
        End If
        
        RaiseEvent AfterUpdate '(Even if unselectable - for navigation.)
    End If
End Property

Public Property Get ValueSelectable() As Boolean
    ValueSelectable = bPValueSelectable
End Property

Public Property Get DayLength() As calMonthLength
    DayLength = lPDayLength
End Property

Public Property Let DayLength(ByVal bDayLength As calMonthLength)
    lPDayLength = bDayLength
    If bInit Then
        Call ApplyWeekDayLabelChanges
    End If
End Property

Public Property Get MonthLength() As calMonthLength
    MonthLength = lPMonthLength
End Property

Public Property Let MonthLength(ByVal iMonthLength As calMonthLength)
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
    Call RenderLabel
End Property


Public Property Get MACFix() As Boolean
    MACFix = bPMACFix
End Property

' MAC Fix
'There is no Transparent buttons in Office in MAC.
'Update: Office 2106 in Windows banned the transparent buttons too.
'
'Normal buttons isn't resizable under a certain size - the text leaning out on the bottom.
'We use labes and transparent buttons for make little size buttons.
'
'This feaure turns off the labels.
Public Property Let MACFix(ByVal bMACFix As Boolean)
    Dim i As Long
    
    bPMACFix = bMACFix
    If bInit Then
        For i = 0 To 41
            mDayButtons(i).Obj_CmBl.Visible = Not bPMACFix
            mDayButtons(i).Obj_CmBlNum.Visible = Not bPMACFix
            mDayButtons(i).Obj_Cmb.Visible = True 'ZOrder
        Next
    End If
    Call Refresh
End Property


Public Property Get RightToLeft() As Boolean
    RightToLeft = bPRightToLeft
End Property

Public Property Let RightToLeft(ByVal bRightToLeft As Boolean)
    bPRightToLeft = bRightToLeft
    If bInit Then
        Call ApplyWeekDayLabelChanges
        Call Refresh
    End If
End Property


'###########################
'# Properties for Day button objects

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

Public Property Get Obj_CmBl() As MSForms.label
    Set Obj_CmBl = CmBl
End Property

Public Property Set Obj_CmBl(ByVal vNewValue As MSForms.label)
    Set CmBl = vNewValue
End Property

Public Property Set Obj_CmBlNum(ByVal vNewValue As MSForms.label)
    Set CmBlNum = vNewValue
End Property

Public Property Get Obj_CmBlNum() As MSForms.label
    Set Obj_CmBlNum = CmBlNum
End Property

Property Let Obj_ValueSelectionEnabled(bSelectable As Boolean)
    If Not mcMain Is Nothing Then
        bPValueSelectable = bSelectable
    End If
End Property

Property Get Obj_ValueSelectionEnabled() As Boolean
    If Not mcMain Is Nothing Then
        Obj_ValueSelectionEnabled = bPValueSelectable
    End If
End Property

'###########################
'# Public Methods

Public Sub AboutBox()
    MsgBox "Calendar Control Class" & vbLf & vbLf & "Autors:" & vbLf & " - r - Original Concept and Base Version" & vbLf & " - Kris - Spirit" & vbLf & " - Gabor - VBA Wizardry and New Features" & vbLf & vbLf & "The FrankensTeam"
End Sub

Public Sub Add(ByVal fForm As MSForms.UserForm)

    Dim cFrame As MSForms.Frame
    Set cFrame = fForm.Controls.Add("Forms.Frame.1")
    
    With cFrame
        .Width = IIf(sPWidth < 0, cDefaultWidth, sPWidth)
        .Height = IIf(sPHeight < 0, cDefaultHeight, sPHeight)
    End With
    
    Call Add_Calendar_into_Frame(cFrame)
    
End Sub

Public Sub Add_Calendar_into_Frame(ByVal cFrame As MSForms.Frame)
    Dim i As Long
    Dim v(199) As Variant
    Dim w As Variant
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
                .BackColor = cDayFontColorSelected 'Dark gray
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
        With mDayButtons(i).Obj_CmBl 'MAC Fix
            .TextAlign = fmTextAlignCenter
            .Visible = Not bPMACFix
        End With
        
        Set mDayButtons(i).Obj_CmBlNum = cFrame.Controls.Add("Forms.Label.1")
        With mDayButtons(i).Obj_CmBlNum
            .TextAlign = fmTextAlignCenter
            .BackStyle = fmBackStyleTransparent
            .Visible = Not bPMACFix
        End With
        
        Set mDayButtons(i).Obj_Cmb = cFrame.Controls.Add("Forms.CommandButton.1")
        With mDayButtons(i).Obj_Cmb
            .BackStyle = fmBackStyleTransparent 'MAC Problem: No button transparency
        End With
        
        mDayButtons(i).RightToLeft = bPRightToLeft
    Next
    
    With CBxY
        .ListRows = 5
        .List = v
        .ListIndex = VBA.Year(dValue) - 1904
        .ShowDropButtonWhen = fmShowDropButtonWhenFocus
        .font.Bold = True
        .MatchRequired = True
    End With

    With CBxM
        .ListRows = 12
        .List = fMonthName(lPMonthLength)
        .ListIndex = VBA.Month(dValue) - 1
        .ShowDropButtonWhen = fmShowDropButtonWhenFocus
        .font.Bold = True
        .MatchRequired = True
    End With
    
    With CLb
        .ForeColor = TitleFontColor
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleTransparent
    End With
    
    Call ApplyWeekDayLabelChanges
    
    Call ApplyFontChanges
    
    Call Refresh_Properities
    
    Call Move
    
End Sub

Private Sub ApplyWeekDayLabelChanges()
    Dim i As Long
    Dim w
    
    w = fWeekdayName(CInt(lPDayLength))
    For i = 0 To 6
        If bPRightToLeft Then
            mLabelButtons(6 - i).Obj_CmBl.Caption = w((i + lPFirstDay - 1) Mod 7)
        Else
            mLabelButtons(i).Obj_CmBl.Caption = w((i + lPFirstDay - 1) Mod 7)
        End If
    Next
End Sub

Private Sub ApplyFontChanges()
    Dim font As MSForms.NewFont
    Dim i As Long

    If Not PDayNewFont Is Nothing Then
        For i = 0 To 6
            Call ApplyFont(mLabelButtons(i).Obj_CmBl.font, DayFont)
        Next
    End If
            
    If Not PGridNewFont Is Nothing Then
        For i = 0 To 41
            If Not bPMACFix Then
                Set font = mDayButtons(i).Obj_CmBlNum.font
            Else
                Set font = mDayButtons(i).Obj_Cmb.font
            End If
            Call ApplyFont(font, GridFont)
        Next
    End If
    
    If Not PTitleNewFont Is Nothing Then
        Call ApplyFont(CLb.font, TitleFont)
    End If

End Sub

Private Sub ApplyFont(fTo As MSForms.NewFont, fFrom As MSForms.NewFont)

    If fTo.Bold <> fFrom.Bold Then _
        fTo.Bold = fFrom.Bold
    If fTo.Weight <> fFrom.Weight Then _
        fTo.Weight = fFrom.Weight
    If fTo.Charset <> fFrom.Charset Then _
        fTo.Charset = fFrom.Charset
    If fTo.Italic <> fFrom.Italic Then _
        fTo.Italic = fFrom.Italic
    If fTo.name <> fFrom.name Then _
        fTo.name = fFrom.name
    If fTo.Size <> fFrom.Size Then _
        fTo.Size = fFrom.Size
    If fTo.Strikethrough <> fFrom.Strikethrough Then _
        fTo.Strikethrough = fFrom.Strikethrough
    If fTo.Underline <> fFrom.Underline Then _
        fTo.Underline = fFrom.Underline

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
    If bInit Then
        Call Refresh_Panel(VBA.Month(dValue), VBA.Year(dValue))
        Call ApplyFontChanges
    End If
End Sub


'###########################
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
    Main.Obj_ValueSelectionEnabled = bPValueSelectable
    Main.Value = dValue
    If bPValueSelectable Then
        Call Main.Event_click
    End If
End Sub

Private Sub CmB_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call Main.Event_DblClick
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
        If bPRightToLeft Then
            newDate = newDate + 1
        Else
            newDate = newDate - 1
        End If
    Case 39
        If bPRightToLeft Then
            newDate = newDate - 1
        Else
            newDate = newDate + 1
        End If
    Case 38
        newDate = newDate - 7
    Case 40
        newDate = newDate + 7
    Case 9
    End Select
    
    If newDate <> dValue Then
        Main.Obj_ValueSelectionEnabled = bPValueSelectable
        Main.Value = newDate
        KeyCode = 0
    Else
        If bPValueSelectable Then
            Call Main.Event_KeyDown(KeyCode, Shift)
        End If
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
    lPSelectedBackColor = &H80000011
    lPUnSelectableBackColor = &H4040C0
    bPVisible = True
    sPHeight = -1
    sPWidth = -1
    sPTop = -1
    sPLeft = -1
    sPControlTipText = ""
    bPRightToLeft = False
    bPSaturdaySelectable = True
    bPSundaySelectable = True
    bPWeekdaySelectable = True
    bPValueSelectable = True
    
    bPMACFix = False
    If val(Application.Version) >= 16 Then '"16.0"
        bPMACFix = True 'Office 2016 compatibility :(
    End If
    
    Set TitleFont = New MSForms.NewFont
    With TitleFont
        .name = "Arial"
        .Size = lPFontSize + 4
        .Bold = True
    End With
    
    Set DayFont = New MSForms.NewFont
    With DayFont
        .name = "Arial"
        .Size = lPFontSize + 2
        .Bold = True
    End With
    
    Set GridFont = New MSForms.NewFont
    With GridFont
        .name = "Arial"
        .Size = lPFontSize
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

Private Function ArraY_Days(ByVal lMonth As Long, ByVal lYear As Long)
    Dim v(0 To 41) As Date, i As Long, g As Long, l As Long, p As Long, t As Date
    
    i = VBA.DateTime.Weekday(VBA.DateSerial(lYear, lMonth, 1), 1 + lPFirstDay Mod 7) - 1
    
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
    
    If bPRightToLeft Then
        For l = 0 To 5
            For i = 0 To 2
                t = v(l * 7 + i)
                v(l * 7 + i) = v(l * 7 + (6 - i))
                v(l * 7 + (6 - i)) = t
            Next
        Next
    End If
    
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
    bInit = (Not CBxY Is Nothing)
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
    Dim m(11), i As Long, v As Variant
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
    Dim m(6), i As Long, v As Variant
    lIndex = lIndex Mod 4
    If Int(lIndex / 2) Then
        If lIndex Mod 2 Then
            v = Array("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
        Else
            v = Array("Monday", "Tuesday", "Wednesday", _
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

Public Sub Event_DblClick()
    RaiseEvent DblClick
End Sub

Public Sub Event_click()
    RaiseEvent Click
End Sub

Public Sub Event_KeyDown( _
    ByVal KeyCode As MSForms.ReturnInteger, _
    ByVal Shift As Integer)
    
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

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

Private Sub Refresh_Selected_Day(ByVal dValue As Date, ByVal i As Long)
    Dim c As MSForms.label
    Dim selColor As OLE_COLOR
    
    If Not bPValueIsNull Then
        bPValueSelectable = mDayButtons(i).Obj_ValueSelectionEnabled
        If bPValueSelectable Then
            selColor = lPSelectedBackColor
        Else
            selColor = lPUnSelectableBackColor
        End If
        On Error Resume Next
        mDayButtons(i).Obj_Cmb.SetFocus
        On Error GoTo 0
        If Not bPMACFix Then
            With mDayButtons(i).Obj_CmBl
                .BackStyle = fmBackStyleOpaque
                .BackColor = selColor
                .ForeColor = cDayFontColorSelected
            End With
        Else
            With mDayButtons(i).Obj_Cmb
                .BackStyle = fmBackStyleOpaque
                .BackColor = selColor
                .ForeColor = cDayFontColorSelected
            End With
        End If
        lPMonth = VBA.Month(dValue)
        lPYear = VBA.Year(dValue)
        lPDay = VBA.Day(dValue)
    End If

End Sub

Private Sub Refresh_Panel(ByVal lMonth As Long, ByVal lYear As Long)
    Dim v As Variant, i As Long, l As Long, idxSel As Variant
    Dim iDay As Long
    Dim lBackColor As OLE_COLOR
    Dim lBackColorA As Variant
    Dim colorArray42() As Variant
    Dim selArray42() As Variant
    Dim bHasColoredDateArray As Boolean
    Dim bSelectable As Boolean
    
    If Not bInit Then
        Exit Sub
    End If
    
    bHasColoredDateArray = HasColoredDateArray()
    
    v = ArraY_Days(lMonth, lYear)
    
    If bHasColoredDateArray Then
        ReDim colorArray42(0 To 41)
        ReDim selArray42(0 To 41)
        
        Call BuildDateColorArrays(colorArray42, selArray42, v(0), v(41))
    End If
    
    idxSel = Empty
    For i = 0 To 41
        mDayButtons(i).Value = v(i)
        If v(i) = dValue Then
            idxSel = i
        End If
        If Not bPMACFix Then 'MAC: no label - command button text
            '# Normal mode
            ' Text day label
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
            ' Day background label
            With mDayButtons(i).Obj_CmBl
                iDay = VBA.DateTime.Weekday(v(i))
                If .BackStyle = fmBackStyleOpaque Then
                    .BackStyle = fmBackStyleTransparent
                End If
                lBackColor = lPBackColor
                If UseDefaultBackColors = False Then
                    If iDay = vbSaturday Then
                        lBackColor = lPSaturdayBackColor
                        If .BackStyle <> fmBackStyleOpaque Then
                            .BackStyle = fmBackStyleOpaque
                        End If
                    ElseIf iDay = vbSunday Then
                        lBackColor = lPSundayBackColor
                        If .BackStyle <> fmBackStyleOpaque Then
                            .BackStyle = fmBackStyleOpaque
                        End If
                    End If
                    If bHasColoredDateArray Then
                        lBackColorA = colorArray42(i)
                        If Not IsEmpty(lBackColorA) Then
                            lBackColor = lBackColorA
                            If .BackStyle <> fmBackStyleOpaque Then
                                .BackStyle = fmBackStyleOpaque
                            End If
                        End If
                    End If
                End If
                If .BackColor <> lBackColor Then
                    .BackColor = lBackColor
                End If
            End With
            ' Button not altered
            With mDayButtons(i).Obj_Cmb
                If .Caption <> "" Then 'After MACFix
                    .Caption = ""
                End If
                If .BackStyle <> fmBackStyleTransparent Then 'Button visible
                    .BackStyle = fmBackStyleTransparent
                End If
            End With
        Else
            '# MAC Fix mode
            With mDayButtons(i).Obj_CmBlNum
                If .Caption <> "" Then
                    .Caption = ""
                End If
            End With
            With mDayButtons(i).Obj_CmBl
                If .BackStyle = fmBackStyleOpaque Then
                    .BackStyle = fmBackStyleTransparent
                End If
            End With
            With mDayButtons(i).Obj_Cmb
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
                If .BackStyle <> fmBackStyleOpaque Then 'Button visible
                    .BackStyle = fmBackStyleOpaque
                End If
                lBackColor = lPBackColor
                If UseDefaultBackColors = False Then
                    iDay = VBA.DateTime.Weekday(v(i))
                    If iDay = vbSaturday Then
                        lBackColor = lPSaturdayBackColor
                    ElseIf iDay = vbSunday Then
                        lBackColor = lPSundayBackColor
                    End If
                    If bHasColoredDateArray Then
                        lBackColorA = colorArray42(i)
                        If Not IsEmpty(lBackColorA) Then
                            lBackColor = lBackColorA
                        End If
                    End If
                End If
                If .BackColor <> lBackColor Then
                    .BackColor = lBackColor
                End If
            End With
        End If
        
        If Not SaturdaySelectable And iDay = vbSaturday Then
            bSelectable = False
        ElseIf Not SundaySelectable And iDay = vbSunday Then
            bSelectable = False
        ElseIf Not WeekdaySelectable And iDay <> vbSaturday And iDay <> vbSunday Then
            bSelectable = False
        Else
            bSelectable = True
        End If
        If bHasColoredDateArray Then
            If Not IsEmpty(selArray42(i)) Then
                bSelectable = selArray42(i)
            End If
        End If
        mDayButtons(i).Obj_ValueSelectionEnabled = bSelectable
        
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
            If mLabelButtons(l).Obj_CmBl.BackStyle = fmBackStyleTransparent Then
                mLabelButtons(l).Obj_CmBl.BackStyle = fmBackStyleOpaque
            End If
            If mLabelButtons(l).Obj_CmBl.BackColor <> lPHeaderBackColor Then
                mLabelButtons(l).Obj_CmBl.BackColor = lPHeaderBackColor
            End If
        Next
    Else
        For l = 0 To 6
            If mLabelButtons(l).Obj_CmBl.BackStyle = fmBackStyleOpaque Then
               mLabelButtons(l).Obj_CmBl.BackStyle = fmBackStyleTransparent
            End If
        Next
    End If
    
    If lMonth = VBA.Month(dValue) And lYear = VBA.Year(dValue) Then
        Call Refresh_Selected_Day(dValue, idxSel)
    Else
        lPMonth = 0
        lPYear = 0
        lPDay = 0
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

'##########################################################
'# Coloring Date Arrays

Public Function HasColoredDateArray() As Boolean
    HasColoredDateArray = Not IsEmpty(maColoredArrayTable)
End Function

Public Function AddColoredDateArray(color As OLE_COLOR, dates As Variant, Optional Selectable As Variant = Empty, Optional index As Long = -1) As Long
    Dim r As Object 'Excel.Range
    Dim dateList() As Variant
    Dim aColoredArrayTable() As Variant
    Dim aColoredArrayRec() As Variant
    Dim newIndex As Long
    Dim lUBnd As Long
    Dim dat As Variant
    Dim format As Integer '1 - 1 dimension, 2 - 1/2 dimension, 3 - 2/2 dimension
    
    If TypeName(dates) = "Variant()" Then
        dateList = dates
    ElseIf TypeName(dates) = "Range" Then
        Set r = dates
        dateList = r.Value2
    Else
        Err.Raise 20001, "Invalid input type for dates: " & TypeName(dates) & " (Valid: Range, Variant())"
    End If
    
    If Not IsEmpty(Selectable) Then
        Selectable = CBool(Selectable)
    End If

    newIndex = index
    If IsEmpty(maColoredArrayTable) Then
        If newIndex = -1 Then
            newIndex = 1
        End If
        ReDim aColoredArrayTable(1 To newIndex)
    Else
        aColoredArrayTable = maColoredArrayTable
        If newIndex = -1 Then
            newIndex = UBound(aColoredArrayTable) + 1
        End If
        If newIndex > UBound(aColoredArrayTable) Then
            ReDim Preserve aColoredArrayTable(1 To newIndex)
        End If
    End If
    
    format = 1
    On Error Resume Next
    lUBnd = UBound(dateList)
    dat = dateList(lUBnd)
    If Err.Number > 0 Then
        Err.Clear
        lUBnd = UBound(dateList, 1)
        dat = dateList(lUBnd, 1)
        If Err.Number > 0 Then
            Err.Raise 20001, "Invalid date array input: " & Err.Description
        End If
        
        format = 2
        If lUBnd < UBound(dateList, 2) Then
            format = 3
        End If
    End If
    On Error GoTo 0
    
    ReDim aColoredArrayRec(1 To 4)
    aColoredArrayRec(ccColor) = color
    aColoredArrayRec(ccFormat) = format
    aColoredArrayRec(ccDateList) = dateList
    aColoredArrayRec(ccSelectable) = Selectable
    
    aColoredArrayTable(newIndex) = aColoredArrayRec
    
    maColoredArrayTable = aColoredArrayTable
    
    Call Refresh
    
    AddColoredDateArray = newIndex
End Function

Public Sub RemoveColoredDateArray(index As Long)
    Dim aColoredArrayTable() As Variant
    Dim i As Long
    Dim bWas As Boolean
    If HasColoredDateArray() Then
        aColoredArrayTable = maColoredArrayTable
        If 1 <= index And index <= UBound(aColoredArrayTable) Then
            aColoredArrayTable(index) = Empty
            bWas = False
            For i = 1 To UBound(aColoredArrayTable)
                If Not IsEmpty(aColoredArrayTable(i)) Then
                   bWas = True
                   Exit For
                End If
            Next
            If bWas Then
                maColoredArrayTable = aColoredArrayTable
            Else
                maColoredArrayTable = Empty
            End If
        End If
    End If
End Sub

Public Sub ClearAllColoredDateArrays()
    maColoredArrayTable = Empty
End Sub


Public Function IsColoredArrayExists(index As Long) As Boolean
    Dim aColoredArrayRec() As Variant
    Call GetColoredArrayRec(index, aColoredArrayRec)
    IsColoredArrayExists = Not IsEmpty(aColoredArrayRec)
End Function

Public Function GetArrayColor(index As Long) As Variant
    Dim aColoredArrayRec() As Variant
    Call GetColoredArrayRec(index, aColoredArrayRec)
    If Not IsEmpty(aColoredArrayRec) Then
        GetArrayColor = aColoredArrayRec(ccColor)
        Exit Function
    End If
    GetArrayColor = Empty
End Function

Public Sub SetArrayColor(index As Long, color As OLE_COLOR)
    Dim aColoredArrayRec() As Variant
    Call GetColoredArrayRec(index, aColoredArrayRec)
    If Not IsEmpty(aColoredArrayRec) Then
        aColoredArrayRec(ccColor) = color
        Call SetColoredArrayRec(index, aColoredArrayRec)
    End If
End Sub

Public Function GetArraySelectable(index As Long) As Variant
    Dim aColoredArrayRec() As Variant
    Call GetColoredArrayRec(index, aColoredArrayRec)
    If Not IsEmpty(aColoredArrayRec) Then
        GetArraySelectable = aColoredArrayRec(ccSelectable)
        Exit Function
    End If
    GetArraySelectable = Empty
End Function

Public Sub SetArraySelectable(index As Long, Selectable As Variant)
    Dim aColoredArrayRec() As Variant
    If Not IsEmpty(Selectable) Then
        Selectable = CBool(Selectable)
    End If
    Call GetColoredArrayRec(index, aColoredArrayRec)
    If Not IsEmpty(aColoredArrayRec) Then
        aColoredArrayRec(ccSelectable) = Selectable
        Call SetColoredArrayRec(index, aColoredArrayRec)
    End If
End Sub

Private Sub GetColoredArrayRec(index As Long, ByRef aColoredArrayRec() As Variant)
    Dim aColoredArrayTable() As Variant
    If HasColoredDateArray() Then
        aColoredArrayTable = maColoredArrayTable
        If 1 <= index And index <= UBound(aColoredArrayTable) Then
            If Not IsEmpty(aColoredArrayTable(index)) Then
                aColoredArrayRec = aColoredArrayTable(index)
                Exit Sub
            End If
        End If
    End If
    aColoredArrayRec = Empty
End Sub

Private Sub SetColoredArrayRec(index As Long, ByRef aColoredArrayRec() As Variant)
    Dim aColoredArrayTable() As Variant
    If HasColoredDateArray() Then
        aColoredArrayTable = maColoredArrayTable
        If 1 <= index And index <= UBound(aColoredArrayTable) Then
            aColoredArrayTable(index) = aColoredArrayRec
            maColoredArrayTable = aColoredArrayTable
        End If
    End If
End Sub


Private Sub BuildDateColorArrays(ByRef colorArray42() As Variant, ByRef selArray42() As Variant, ByVal fromDate As Date, ByVal toDate As Date)
    Dim aColoredArrayTable() As Variant
    Dim aColoredArrayRec() As Variant
    Dim iDate As Date
    Dim format As Integer '1 - 1 dimension, 2 - 2 dimension/1, 2 - 2 dimension/2
    Dim dateList() As Variant
    Dim i As Long
    Dim j As Long
    Dim idx As Integer

    If Not HasColoredDateArray() Then
        Exit Sub
    End If
    
    aColoredArrayTable = maColoredArrayTable
    
    For i = 1 To UBound(aColoredArrayTable)
        aColoredArrayRec = aColoredArrayTable(i)
        
        format = aColoredArrayRec(ccFormat)
        dateList = aColoredArrayRec(ccDateList)
        
        Select Case format
        Case 1
            For j = LBound(dateList) To UBound(dateList)
                iDate = dateList(j)
                If fromDate <= iDate And iDate <= toDate Then
                    idx = iDate - fromDate
                    colorArray42(idx) = aColoredArrayRec(ccColor)
                    selArray42(idx) = aColoredArrayRec(ccSelectable)
                End If
            Next
        Case 2
            For j = LBound(dateList, 1) To UBound(dateList, 1)
                iDate = dateList(j, 1)
                If fromDate <= iDate And iDate <= toDate Then
                    idx = iDate - fromDate
                    colorArray42(idx) = aColoredArrayRec(ccColor)
                    selArray42(idx) = aColoredArrayRec(ccSelectable)
                End If
            Next
        Case 3
            For j = LBound(dateList, 2) To UBound(dateList, 2)
                iDate = dateList(1, j)
                If fromDate <= iDate And iDate <= toDate Then
                    idx = iDate - fromDate
                    colorArray42(idx) = aColoredArrayRec(ccColor)
                    selArray42(idx) = aColoredArrayRec(ccSelectable)
                End If
            Next
        End Select
    Next

End Sub

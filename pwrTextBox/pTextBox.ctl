VERSION 5.00
Begin VB.UserControl pTextBox 
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1575
   ScaleHeight     =   285
   ScaleWidth      =   1575
   ToolboxBitmap   =   "pTextBox.ctx":0000
   Begin VB.TextBox pwrtbox 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   0
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "pTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum afAppearance
    pFlat = 0
    p3D = 1
End Enum
Public Enum afAlignment
    pLeft = 0
    pRight = 1
    pCenter = 2
End Enum
Public Enum afBorderStyle
    pNone = 0
    pFixed = 1
End Enum
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Change()
Event Click()
Event DblClick()
Event Resize()

Dim oSubClass As clsSubClass

Public Property Get Alignment() As afAlignment
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = pwrtbox.Alignment
End Property
Public Property Let Alignment(ByVal New_Alignment As afAlignment)
    pwrtbox.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property
Public Property Get Appearance() As afAppearance
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = pwrtbox.Appearance
End Property
Public Property Let Appearance(ByVal New_Appearance As afAppearance)
    pwrtbox.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = pwrtbox.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    pwrtbox.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property
Public Property Get BorderStyle() As afBorderStyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = pwrtbox.BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As afBorderStyle)
    pwrtbox.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property
Public Property Get CausesValidation() As Boolean
Attribute CausesValidation.VB_Description = "Returns/sets whether validation occurs on the control which lost focus."
    CausesValidation = pwrtbox.CausesValidation
End Property
Public Property Let CausesValidation(ByVal New_CausesValidation As Boolean)
    pwrtbox.CausesValidation() = New_CausesValidation
    PropertyChanged "CausesValidation"
End Property
Private Sub pwrtBox_Change()
    RaiseEvent Change
End Sub
Private Sub pwrtBox_Click()
    RaiseEvent Click
End Sub
Private Sub pwrtBox_DblClick()
    RaiseEvent DblClick
End Sub
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = pwrtbox.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    pwrtbox.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = pwrtbox.Font
End Property
Public Sub Refresh()
    UpdateWindow UserControl.hWnd
End Sub

Public Property Set Font(ByVal New_Font As Font)
    Set pwrtbox.Font = New_Font
    PropertyChanged "Font"
End Property
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = pwrtbox.ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    pwrtbox.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = pwrtbox.Locked
End Property
Public Property Let Locked(ByVal New_Locked As Boolean)
    pwrtbox.Locked() = New_Locked
    PropertyChanged "Locked"
End Property
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
    MaxLength = pwrtbox.MaxLength
End Property
Public Property Let MaxLength(ByVal New_MaxLength As Long)
    pwrtbox.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = pwrtbox.MouseIcon
End Property
Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set pwrtbox.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = pwrtbox.MousePointer
End Property
Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    pwrtbox.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get PasswordChar() As String
    PasswordChar = pwrtbox.PasswordChar
End Property
Public Property Let PasswordChar(ByVal New_PasswordChar As String)

   pwrtbox.PasswordChar() = New_PasswordChar
   PropertyChanged "PasswordChar"
End Property

Private Sub UserControl_Initialize()
 'Create our new subclassing object
    Set oSubClass = New clsSubClass
    oSubClass.hWnd = UserControl.pwrtbox.hWnd
    'Start Subclassing.
    oSubClass.Attach
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    With UserControl
        .pwrtbox.Move 0, 0, .ScaleWidth, .ScaleHeight
        If .Height < .pwrtbox.Height Then .Height = .pwrtbox.Height
    End With
    RaiseEvent Resize
End Sub
Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "Determines text display direction and control visual appearance on a bidirectional system."
    RightToLeft = pwrtbox.RightToLeft
End Property
Public Property Let RightToLeft(ByVal New_RightToLeft As Boolean)
    pwrtbox.RightToLeft() = New_RightToLeft
    PropertyChanged "RightToLeft"
End Property
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
    SelLength = pwrtbox.SelLength
End Property
Public Property Let SelLength(ByVal New_SelLength As Long)
    pwrtbox.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
    SelStart = pwrtbox.SelStart
End Property
Public Property Let SelStart(ByVal New_SelStart As Long)
    pwrtbox.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
    SelText = pwrtbox.SelText
End Property
Public Property Let SelText(ByVal New_SelText As String)
    pwrtbox.SelText() = New_SelText
    PropertyChanged "SelText"
End Property
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = pwrtbox.ToolTipText
End Property
Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    pwrtbox.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    pwrtbox.Alignment = PropBag.ReadProperty("Alignment", 0)
    pwrtbox.Appearance = PropBag.ReadProperty("Appearance", 1)
    pwrtbox.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    pwrtbox.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    pwrtbox.CausesValidation = PropBag.ReadProperty("CausesValidation", False)
    pwrtbox.Enabled = PropBag.ReadProperty("Enabled", True)
    Set pwrtbox.Font = PropBag.ReadProperty("Font", Ambient.Font)
    pwrtbox.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    pwrtbox.Locked = PropBag.ReadProperty("Locked", False)
    pwrtbox.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    pwrtbox.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    pwrtbox.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    pwrtbox.RightToLeft = PropBag.ReadProperty("RightToLeft", False)
    pwrtbox.SelLength = PropBag.ReadProperty("SelLength", 0)
    pwrtbox.SelStart = PropBag.ReadProperty("SelStart", 0)
    pwrtbox.SelText = PropBag.ReadProperty("SelText", "")
    pwrtbox.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    pwrtbox.Text = PropBag.ReadProperty("Text", "Text1")
    
End Sub

Private Sub UserControl_Terminate()
   oSubClass.Detach
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Alignment", pwrtbox.Alignment, 0)
    Call PropBag.WriteProperty("Appearance", pwrtbox.Appearance, 1)
    Call PropBag.WriteProperty("BackColor", pwrtbox.BackColor, &H80000005)
    Call PropBag.WriteProperty("BorderStyle", pwrtbox.BorderStyle, 1)
    Call PropBag.WriteProperty("CausesValidation", pwrtbox.CausesValidation, False)
    Call PropBag.WriteProperty("Enabled", pwrtbox.Enabled, True)
    Call PropBag.WriteProperty("Font", pwrtbox.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", pwrtbox.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Locked", pwrtbox.Locked, False)
    Call PropBag.WriteProperty("MaxLength", pwrtbox.MaxLength, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", pwrtbox.MousePointer, 0)
    Call PropBag.WriteProperty("PasswordChar", pwrtbox.PasswordChar, "")
    Call PropBag.WriteProperty("RightToLeft", pwrtbox.RightToLeft, False)
    Call PropBag.WriteProperty("SelLength", pwrtbox.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", pwrtbox.SelStart, 0)
    Call PropBag.WriteProperty("SelText", pwrtbox.SelText, "")
    Call PropBag.WriteProperty("ToolTipText", pwrtbox.ToolTipText, "")
    Call PropBag.WriteProperty("Text", pwrtbox.Text, "Text1")
    
End Sub
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = pwrtbox.Text
End Property
Public Property Let Text(ByVal New_Text As String)
    pwrtbox.Text() = New_Text
    PropertyChanged "Text"
End Property
Private Sub pwrtBox_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub pwrtBox_KeyPress(KeyAscii As Integer)
    
    RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub pwrtBox_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

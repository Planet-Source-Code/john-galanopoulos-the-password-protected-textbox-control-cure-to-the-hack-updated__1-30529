VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tester"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      TabIndex        =   15
      Top             =   6840
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4200
      TabIndex        =   14
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">>>"
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">>>"
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4200
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">>>"
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4200
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "myhiddenpass"
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>>"
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "John Galanopoulos  <GreekThought@Yahoo.gr>"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   6840
      Width           =   3855
   End
   Begin VB.Label Label9 
      Caption         =   $"Form1.frx":0000
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   5640
      Width           =   8775
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   6600
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label8 
      Caption         =   "The VB classic  textbox"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   1815
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   6600
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   6600
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label7 
      Caption         =   "Our subclassed textbox"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Result"
      Height          =   255
      Left            =   6240
      TabIndex        =   11
      Top             =   4080
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6600
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label5 
      Caption         =   $"Form1.frx":01C7
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Method 1 : Send a message to the text box using WM_GETTEXT"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Result"
      Height          =   255
      Left            =   6120
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":024F
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Method 1 : Send a message to the protected text box using WM_GETTEXT"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
   ByVal lpParameters As String, ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long
   
Private Const EM_SETPASSWORDCHAR = &HCC
Private Const WM_GETTEXT = &HD
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Private Sub Command1_Click()
Dim hwndTextBox As Long
Dim revpass As String
Dim smRet As Long
Dim lngpass As Long

lngpass = 256
revpass = Space$(lngpass)

hwndTextBox = Text1.hWnd

 smRet = SendMessage(hwndTextBox, WM_GETTEXT, lngpass, ByVal revpass)
 revpass = Trim$(revpass)
      
          If smRet = 0 Then
                 Text2.Text = "N/A"
            ElseIf revpass = vbNullString Then
                 Text2.Text = "N/A"
            Else
                 Text2.Text = revpass
          End If
End Sub

Private Sub Command2_Click()
Dim hwndTextBox As Long
Dim smRet As Long

hwndTextBox = Text1.hWnd

smRet = SendMessage(hwndTextBox, EM_SETPASSWORDCHAR, ByVal 0, 0)
Text1.Refresh

End Sub

Private Sub Command3_Click()
Dim hwndTextBox As Long
Dim revpass As String
Dim smRet As Long
Dim lngpass As Long

lngpass = 256
revpass = Space$(lngpass)

hwndTextBox = pTextBox1.hWnd

 smRet = SendMessage(hwndTextBox, WM_GETTEXT, lngpass, ByVal revpass)
 revpass = Trim$(revpass)
      
          If smRet = 0 Then
                 Text3.Text = "N/A"
            ElseIf revpass = vbNullString Then
                 Text3.Text = "N/A"
            Else
                 Text3.Text = revpass
          End If
End Sub

Private Sub Command4_Click()
Dim hwndTextBox As Long
Dim smRet As Long

hwndTextBox = pTextBox1.hWnd

smRet = SendMessage(hwndTextBox, EM_SETPASSWORDCHAR, ByVal 0, 0)
pTextBox1.Refresh

End Sub


Private Sub Command5_Click()
If MsgBox("That's all folks. Hope u liked it. Any comments, suggestions or votes are always welcomed. Although voting or commenting is not necessary, it is a good way to support the contribution to knowledge, by this coder and many more on PCS. Comments or suggestions are always very helpful cause it gives us the chance to improve our source. Anyway, if you want to navigate to the site you downloaded this source" & _
           vbCrLf & "and submit your comment or vote click OK. " & vbCrLf & "Happy, wealthy, healthy new year to all.", vbInformation + vbOKCancel, Year(Now())) = vbOK Then
           Dim vlink As String
           vlink = "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&blnAuthorSearch=TRUE&lngAuthorId=206793&strAuthorName=John%20Galanopoulos&txtMaxNumberOfEntriesPerPage=25"
           ShellExecute 0, vbNullString, vlink, vbNullString, vbNullString, vbNormalFocus
End If
End Sub

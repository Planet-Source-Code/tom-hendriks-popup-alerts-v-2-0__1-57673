VERSION 5.00
Begin VB.Form frmGmail 
   BorderStyle     =   0  'None
   ClientHeight    =   2010
   ClientLeft      =   5985
   ClientTop       =   2715
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3840
      Top             =   720
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   7000
      Left            =   3360
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2880
      Top             =   720
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   120
      Picture         =   "frmGmail.frx":0000
      Top             =   120
      Width           =   585
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmGmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'All the API calls here are for form transparency :-)

'This is the simple one which only works in Win XP/2k,
'so other users will get a normal grey form :-(.

Private Declare Function GetWindowLong Lib "user32" Alias _
    "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long


Private Declare Function SetWindowLong Lib "user32" Alias _
    "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long


Private Declare Function SetLayeredWindowAttributes Lib "user32" _
    (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, _
    ByVal dwFlags As Long) As Long
    
    Private Const Style = (-20)
    Private Const NewLong = &H80000
    Private Const Alpha = &H2&
   
   
   Dim PosX As Long
   Dim PosY As Long

Private Sub Form_Click()
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = True

End Sub

Private Sub Form_Load()

    Dim lForm As Long
    Dim Transparency As Byte ' Transparency (0 - 255)
    Transparency = 180
    lForm = GetWindowLong(Me.hwnd, Style)
    SetWindowLong Me.hwnd, Style, NewLong
    SetLayeredWindowAttributes Me.hwnd, 0, Transparency, Alpha
    
    PosX = Screen.Width - Me.Width
PosY = Screen.Height

Me.Left = PosX
Me.Top = PosY

End Sub

Private Sub lblText_Click()
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = True
End Sub

Private Sub lblTitle_Click()
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = True
End Sub

Private Sub Timer1_Timer()
Me.Top = Me.Top - 30
If Me.Top < PosY - Me.Height Then
    Timer1.Enabled = False
    Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False
Timer3.Enabled = True
End Sub

Private Sub Timer3_Timer()
Me.Top = Me.Top + 30
If Me.Top = PosY Then
    Timer3.Enabled = False
    Unload Me
End If

End Sub

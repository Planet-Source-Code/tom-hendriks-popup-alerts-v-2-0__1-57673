VERSION 5.00
Begin VB.Form frmNewAlert 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   ClientHeight    =   1815
   ClientLeft      =   12495
   ClientTop       =   11520
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1560
      Top             =   720
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   960
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   360
      Top             =   720
   End
   Begin VB.Shape Shape1 
      Height          =   1815
      Left            =   0
      Top             =   0
      Width           =   2775
   End
   Begin VB.Image imgPointer 
      Height          =   480
      Left            =   1920
      Picture         =   "frmAlert3.frx":0000
      Top             =   3600
      Width           =   480
   End
   Begin VB.Label lblURL 
      Caption         =   "Label1"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label lblLink 
      Caption         =   "Label1"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Image imgXNormal 
      Height          =   240
      Left            =   960
      Picture         =   "frmAlert3.frx":030A
      Top             =   2520
      Width           =   225
   End
   Begin VB.Image imgXOver 
      Height          =   225
      Left            =   720
      Picture         =   "frmAlert3.frx":064C
      Top             =   2520
      Width           =   210
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Left            =   120
      TabIndex        =   1
      Top             =   75
      Width           =   2535
   End
   Begin VB.Label lblAlert 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Your text here"
      Height          =   1095
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "frmNewAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim PosX As Long
Dim PosY As Long




Private Sub Form_Load()

PosX = Screen.Width - Me.Width
PosY = Screen.Height

Me.Left = PosX
Me.Top = PosY
End Sub



Private Sub lblAlert_Click()
If lblLink.Caption = "True" Then
    Call ShellExecute(Me.hwnd, "Open", lblURL.Caption, 0, 0, 10)

End If
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

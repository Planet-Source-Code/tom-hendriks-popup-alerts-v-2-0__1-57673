VERSION 5.00
Begin VB.Form frmAlertOld 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2190
   ClientLeft      =   11325
   ClientTop       =   8640
   ClientWidth     =   2550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   2550
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2115
      ScaleWidth      =   2475
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1440
         Top             =   1440
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   7000
         Left            =   840
         Top             =   1440
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   240
         Top             =   1440
      End
      Begin VB.Label lblMessage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00400000&
         Height          =   1215
         Left            =   120
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   600
         Width           =   2295
      End
      Begin VB.Image imgX 
         Height          =   225
         Left            =   2160
         Picture         =   "frmAlert2.frx":0000
         Top             =   120
         Width           =   210
      End
      Begin VB.Image frmAlert2 
         Height          =   2115
         Left            =   0
         Picture         =   "frmAlert2.frx":02D6
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2715
      End
   End
   Begin VB.Label lblLink 
      Caption         =   "Label1"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label lblurl 
      Caption         =   "Label1"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Image imgXNormal 
      Height          =   225
      Left            =   1800
      Picture         =   "frmAlert2.frx":0A0D
      Top             =   4200
      Width           =   210
   End
   Begin VB.Image imgPointer 
      Height          =   480
      Left            =   720
      Picture         =   "frmAlert2.frx":0CE3
      Top             =   3840
      Width           =   480
   End
   Begin VB.Image imgXOver 
      Height          =   225
      Left            =   1080
      Picture         =   "frmAlert2.frx":0FED
      Top             =   2880
      Width           =   210
   End
End
Attribute VB_Name = "frmAlertOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim PosX As Long
Dim PosY As Long



Private Sub Form_Load()
If lblLink.Caption = "True" Then
    lblMessage.ForeColor = vbBlue

End If

PosX = Screen.Width - Me.Width
PosY = Screen.Height

Me.Left = PosX
Me.Top = PosY
End Sub

Private Sub frmAlert2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgX.Picture = imgXNormal.Picture
End Sub





Private Sub imgX_Click()
Unload Me
End Sub

Private Sub imgX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgX.Picture = imgXOver.Picture
End Sub

Private Sub lblMessage_Click()
If lblLink.Caption = "True" Then

    Call ShellExecute(Me.hwnd, "Open", lblURL.Caption, 0, 0, 10)

End If
End Sub

Private Sub Timer1_Timer()
Me.Top = Me.Top - 40

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
Me.Top = Me.Top + 40

If Me.Top = PosY Then
    Timer3.Enabled = False
    Unload Me

End If
End Sub

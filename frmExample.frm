VERSION 5.00
Begin VB.Form frmExample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Example"
   ClientHeight    =   2715
   ClientLeft      =   4890
   ClientTop       =   3795
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   4200
   Begin VB.CommandButton Command4 
      Caption         =   "Gmail Notifier"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Flat Style"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Old MSN"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MSN Version 6.2"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmExample.frx":0000
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   3975
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'Set variables for popup

Dim Newalert As New cAlert
Newalert.Message = "Testing..."
Newalert.Title = "MSN 6.2 Style"
Newalert.Link = True
Newalert.sUrl = "http://www.google.com"
Newalert.MSN6



End Sub

Private Sub Command2_Click()
'Set variables for popup

Dim Newalert As New cAlert
Newalert.Link = True
Newalert.sUrl = "http://www.google.com"
Newalert.Message = "Old MSN style test..."
Newalert.MSNOld
End Sub

Private Sub Command3_Click()
Dim Newalert As New cAlert
Newalert.Title = "New Style Test"
Newalert.Message = "Message Here"
Newalert.FlatStyle
End Sub

Private Sub Command4_Click()

Dim Newalert As New cAlert
Newalert.Title = "Gmail Notifier Style Test"
Newalert.Message = "Message Here"
Newalert.Gmail

End Sub

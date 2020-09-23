VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------
'***************************************************
'*******************Popup Alerts********************
'**************(c) 2003-4 Tom Hendriks**************
'***************************************************
'|**********************************************|
'|Version history:                              |
'|                                              |
'|v1 - MSN 6 style popup coded                  |
'|v1.1 - Old MSN and Flat styles added          |
'|----------------------------------------------|
'|v2 - Gmail Notifier style added               |
'|     New positioning to ensure correct        |
'|     display at all resolutions               |
'|**********************************************|




Option Explicit
Public Title As String
Public Message As String
Public Link As Boolean
Public sUrl As String
Public PosX As Long
Public PosY As Long


Public Sub MSN6()

'Newest MSN style popup

'New Alert
Dim alert As New frmAlert6




'Set the Title and Message
alert.lblTitle.Caption = Title
alert.lblAlert.Caption = Message

'If link option is used then use "url hand" icon
'and tell the alert where to link to.

If Link = True Then
alert.lblAlert.MouseIcon = alert.imgPointer.Picture
alert.lblURL.Caption = sUrl
alert.lblLink.Caption = "True"
End If

'Display the alert

alert.Show





End Sub

Public Sub FlatStyle()

'New Alert by me
'Doesn't support links

'New Alert
Dim alert As New frmNewAlert


'Set the Title and Message
alert.lblTitle.Caption = Title
alert.lblAlert.Caption = Message


'Display the alert

alert.Show





End Sub

Public Sub Gmail()

'Clone of the alerts from the Gmail Notifier app

Dim alert As New frmGmail



alert.lblTitle.Caption = Title
alert.lblText.Caption = Message


alert.Show

End Sub


Public Sub MSNOld()

'New Alert
Dim alert As New frmAlertOld


'Set Message (no title used here)
alert.lblMessage.Caption = Message

'If link option is used then use "url hand" icon
'and tell the alert where to link to.
If Link = True Then
alert.lblMessage.MouseIcon = alert.imgPointer.Picture
alert.lblURL.Caption = sUrl
alert.lblLink.Caption = "True"

End If

'Display the alert
alert.Show


End Sub

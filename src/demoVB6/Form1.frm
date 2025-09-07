VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   960
      Top             =   2280
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   2640
      TabIndex        =   0
      Top             =   2040
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim t As New cThread
Dim WithEvents ui As cThreadUINotifier
Attribute ui.VB_VarHelpID = -1

Private Sub Form_Load()
    t.Create AddressOf Module1.ThreadProc
End Sub

Private Sub Timer1_Timer()
    Me.Caption = VBMAN.Json.RootItems.Count
End Sub

Private Sub ui_ThreadNotification(ByVal TaskId As String, ByVal NotificationCode As Long, ByVal Data As Variant)
    
End Sub

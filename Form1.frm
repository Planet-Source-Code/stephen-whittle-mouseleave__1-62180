VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3165
   LinkTopic       =   "Form1"
   ScaleHeight     =   1995
   ScaleWidth      =   3165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Not Over"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long


Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

 If Command1.Caption = "Not Over" Then
    SetCapture Command1.hwnd
    Command1.Caption = "Now Over"
 ElseIf X < 0 Or X > Command1.Width Or Y < 0 Or Y > Command1.Height Then
    Command1.Caption = "Not Over"
    ReleaseCapture
 End If

End Sub



VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Text            =   "100000"
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Record"
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Don't forget to visit my web site at http://vbplace.htmlplanet.com"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":0000
      Height          =   1215
      Left            =   2520
      TabIndex        =   4
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Times"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   960
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Dim Position As POINTAPI
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Private Sub Command1_Click()
Command1.Enabled = False
Kill "c:\tempX.tmp"
Kill "c:\tempY.tmp"
Open "c:\tempX.tmp" For Output As #1
Open "c:\tempY.tmp" For Output As #2
i = 0
Do
    GetCursorPos Position
    Print #1, Position.X
    Print #2, Position.Y
    i = i + 1
Loop Until i = Text1.Text
Close #1: Close #2
Command1.Enabled = True
Command2.Caption = "Play " + Text1.Text + " times"
End Sub

Private Sub Command2_Click()
Command2.Enabled = False
Open "c:\tempX.tmp" For Input As #1
Open "c:\tempY.tmp" For Input As #2

i = 0
Do
    Input #1, X
    Input #2, Y
    SetCursorPos X, Y
i = i + 1
Loop Until i = Text1.Text
Close #1: Close #2
Command2.Enabled = True
End Sub


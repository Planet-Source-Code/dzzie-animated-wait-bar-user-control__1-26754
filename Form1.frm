VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Display Bar User Control Demo"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin Project1.dBar dbar1 
      Height          =   360
      Left            =   75
      TabIndex        =   5
      Top             =   45
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   635
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Pulsate"
      Height          =   300
      Index           =   3
      Left            =   1560
      TabIndex        =   4
      Top             =   975
      Width           =   1005
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Bounce"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   975
      Width           =   1170
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Double Bounce"
      Height          =   300
      Index           =   2
      Left            =   1560
      TabIndex        =   2
      Top             =   675
      Width           =   1485
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Left to Right"
      Height          =   285
      Index           =   0
      Left            =   225
      TabIndex        =   1
      Top             =   660
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Start"
      Height          =   375
      Left            =   3675
      TabIndex        =   0
      Top             =   900
      Width           =   1110
   End
   Begin VB.Label Label1 
      Caption         =   "Note:  Max Resize Height of Control = 315, Max Width = Unlimited"
      Height          =   225
      Left            =   90
      TabIndex        =   6
      Top             =   1410
      Width           =   4845
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'simple user control to keep users occupied in times of wait
'adds about 16k to your compiled exe

'this example uses the default timing and interval
'properties, you can set both manually if desired
'dbar1.Timimg = 100
'dbar1.Increment = 200
    
Private Sub Option1_Click(Index As Integer)
    
    If cmd.Caption = "Stop" Then
        cmd.Caption = "Start"
        dbar1.EndDisplay
        Call cmd_Click
    End If

End Sub

Private Sub cmd_Click()

    If cmd.Caption = "Stop" Then
        dbar1.EndDisplay
        cmd.Caption = "Start"
    Else
        If Option1(0).Value Then dbar1.Style = Monic
        If Option1(1).Value Then dbar1.Style = Pacer
        If Option1(2).Value Then dbar1.Style = Bouncy
        If Option1(3).Value Then dbar1.Style = Pulse
        dbar1.BeginDisplay
        cmd.Caption = "Stop"
    End If

End Sub



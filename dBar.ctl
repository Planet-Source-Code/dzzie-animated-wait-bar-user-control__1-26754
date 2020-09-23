VERSION 5.00
Begin VB.UserControl dBar 
   ClientHeight    =   990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   990
   ScaleWidth      =   4800
   ToolboxBitmap   =   "dBar.ctx":0000
   Begin VB.Timer tmrPulse 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2925
      Top             =   480
   End
   Begin VB.Timer tmrMonic 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1425
      Top             =   480
   End
   Begin VB.Timer tmrPace 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1905
      Top             =   480
   End
   Begin VB.Timer tmrBouncy 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2415
      Top             =   480
   End
   Begin VB.PictureBox pBar 
      Height          =   315
      Left            =   150
      ScaleHeight     =   255
      ScaleWidth      =   4365
      TabIndex        =   0
      Top             =   150
      Width           =   4425
      Begin VB.PictureBox pLeft 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   0
         Picture         =   "dBar.ctx":0312
         ScaleHeight     =   300
         ScaleWidth      =   750
         TabIndex        =   2
         Top             =   0
         Width           =   750
      End
      Begin VB.PictureBox pRight 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   3615
         Picture         =   "dBar.ctx":0F34
         ScaleHeight     =   300
         ScaleWidth      =   750
         TabIndex        =   1
         Top             =   -15
         Width           =   750
      End
      Begin VB.Label lblPulse 
         BackColor       =   &H00808000&
         Height          =   285
         Left            =   1155
         TabIndex        =   3
         Top             =   -15
         Width           =   2010
      End
   End
End
Attribute VB_Name = "dBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type flags
    STY As DisplayStyle
    LOS As Long    'where Left goes completly Off Screen
    ROS As Long    'where Right goes completly Off Screen
    INC As Integer 'INCrement amount
    TMR As Integer 'timer interval
    EXP As Boolean 'EXPand stage for pulse
End Type

Public Enum DisplayStyle
    Monic = 1
    Pacer = 2
    Bouncy = 3
    Pulse = 4
End Enum

Dim f As flags


Property Let Style(s As DisplayStyle)
  f.STY = s
End Property

Property Let Timimg(milsec As Integer)
  f.TMR = milsec
End Property

Property Let Increment(twips As Integer)
  f.INC = twips
End Property



Private Sub UserControl_Resize()
    pBar.Width = UserControl.Width - 200
    pBar.Height = UserControl.Height - 200
    If pBar.Height < 315 Then pBar.Height = UserControl.Height - 200 _
    Else pBar.Height = 315
    Call UserControl_Initialize
End Sub

Private Sub UserControl_Initialize()
    f.EXP = True
    tmrPace.Enabled = False
    tmrMonic.Enabled = False
    tmrBouncy.Enabled = False
    tmrPulse.Enabled = False
    lblPulse.Visible = False
    pLeft.Left = pBar.Left - pLeft.Width - 150
    pRight.Left = pBar.Left + pBar.Width - 150
    f.LOS = pLeft.Left
    f.ROS = pRight.Left
End Sub


Sub EndDisplay()
    Call UserControl_Initialize
End Sub

Sub BeginDisplay()
   Call UserControl_Initialize
   
   'set defaults
   If f.TMR = 0 Then f.TMR = 50
   If f.INC = 0 Then f.INC = 200
   If f.STY = Empty Then f.STY = Monic
   
   If f.STY = Monic Then
        tmrMonic.Interval = f.TMR
        tmrMonic.Enabled = True
   ElseIf f.STY = Pacer Then
        tmrPace.Interval = f.TMR
        tmrPace.Enabled = True
   ElseIf f.STY = Bouncy Then
        tmrBouncy.Interval = f.TMR
        tmrBouncy.Enabled = True
   Else
        lblPulse.Visible = True
        lblPulse.Width = 0
        tmrPulse.Interval = f.TMR
        tmrPulse.Enabled = True
   End If
End Sub



Private Sub tmrMonic_Timer()
    pLeft.Left = pLeft.Left + f.INC
    If pLeft.Left >= f.ROS Then pLeft.Left = f.LOS
End Sub

Private Sub tmrPace_Timer()
    If pLeft.Left < f.ROS Then
        pLeft.Left = pLeft.Left + f.INC
    ElseIf pLeft.Left >= f.ROS And pRight.Left > f.LOS Then
        pRight.Left = pRight.Left - f.INC
    Else
         pLeft.Left = f.LOS
         pRight.Left = f.ROS
    End If
End Sub

Private Sub tmrBouncy_Timer()
    If pLeft.Left < f.ROS Then
        pLeft.Left = pLeft.Left + f.INC
    Else
        pLeft.Left = f.LOS
    End If
    
    If pRight.Left > f.LOS Then
        pRight.Left = pRight.Left - f.INC
    Else
        pRight.Left = f.ROS
    End If
End Sub

Private Sub tmrPulse_Timer()
    If f.EXP = True Then
        lblPulse.Width = lblPulse.Width + f.INC
        lblPulse.Left = pBar.Left + (pBar.Width / 2) - (lblPulse.Width / 2) - 300
        If lblPulse.Width > pBar.Width Then f.EXP = False
    Else
        If lblPulse.Width - f.INC < 0 Then
            f.EXP = True: Exit Sub
        Else
            lblPulse.Width = lblPulse.Width - f.INC
            lblPulse.Left = pBar.Left + (pBar.Width / 2) - (lblPulse.Width / 2) - 300
        End If
    End If
End Sub

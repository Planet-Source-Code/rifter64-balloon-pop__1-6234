VERSION 5.00
Begin VB.Form frmBalloonPop 
   BackColor       =   &H00000040&
   Caption         =   "The Balloon Pop"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8145
   ControlBox      =   0   'False
   Icon            =   "Balloon Pop.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "EXIT GAME"
      Height          =   735
      Left            =   1680
      TabIndex        =   14
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Timer TmrTime2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3000
      Top             =   6000
   End
   Begin VB.Timer TmrTime1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Top             =   6000
   End
   Begin VB.Timer tmrpop2 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   600
      Top             =   6000
   End
   Begin VB.Timer TmrPop 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   120
      Top             =   6000
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "START GAME"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   1575
   End
   Begin VB.PictureBox pc1 
      Height          =   1815
      Left            =   240
      ScaleHeight     =   1755
      ScaleWidth      =   1035
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
   Begin VB.PictureBox pc2 
      Height          =   1815
      Left            =   1560
      ScaleHeight     =   1755
      ScaleWidth      =   1035
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.PictureBox pc3 
      Height          =   1815
      Left            =   2880
      ScaleHeight     =   1755
      ScaleWidth      =   1035
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.PictureBox pc4 
      Height          =   1815
      Left            =   4200
      ScaleHeight     =   1755
      ScaleWidth      =   1035
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.PictureBox pc5 
      Height          =   1815
      Left            =   5520
      ScaleHeight     =   1755
      ScaleWidth      =   1035
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.PictureBox pc6 
      Height          =   1815
      Left            =   6840
      ScaleHeight     =   1755
      ScaleWidth      =   1035
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   8040
      X2              =   120
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   8040
      X2              =   8040
      Y1              =   600
      Y2              =   3000
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   600
      Y2              =   3000
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   120
      X2              =   8040
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   3840
      X2              =   7920
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   7920
      X2              =   7920
      Y1              =   4800
      Y2              =   3240
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   3840
      X2              =   7920
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   3840
      X2              =   7920
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   3840
      X2              =   3840
      Y1              =   3240
      Y2              =   4800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   3840
      X2              =   7920
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      Caption         =   "Pop As Many Balloons As You Can In 2 Minutes."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Width           =   3255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   0
      X2              =   8160
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label lblover 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      Caption         =   "GAME OVER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.Label lblFinish 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3960
      TabIndex        =   11
      Top             =   4200
      Width           =   3855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      Caption         =   "YOU'RE SCORE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label lblTimeLeft 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6840
      TabIndex        =   9
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      Caption         =   "TIME LEFT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Label lblCount 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6840
      TabIndex        =   7
      Top             =   3360
      Width           =   975
   End
   Begin VB.Image ImgBlank 
      Height          =   1800
      Left            =   2520
      Picture         =   "Balloon Pop.frx":030A
      Top             =   4920
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image ImgBall 
      Height          =   1800
      Left            =   1320
      Picture         =   "Balloon Pop.frx":0629
      Top             =   4920
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image ImgPop 
      Height          =   1800
      Left            =   120
      Picture         =   "Balloon Pop.frx":0C7E
      Top             =   4920
      Visible         =   0   'False
      Width           =   1035
   End
End
Attribute VB_Name = "frmBalloonPop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub DisableCtrlAltDelete(bDisabled As Boolean)
    Dim X As Long
    X = SystemParametersInfo(97, bDisabled, CStr(1), 0)
End Sub
    Private Sub Pop_Balloon()
        iTalk = sndPlaySound(ByVal CStr("c:\myshit~1\balloo~1\pop.wav"), Sync)
End Sub

Private Sub cmdExit_Click()
    Call DisableCtrlAltDelete(False)
    End
End Sub

Private Sub cmdStart_Click()
    Let pc1.Picture = ImgBlank
    Let pc2.Picture = ImgBlank
    Let pc3.Picture = ImgBlank
    Let pc4.Picture = ImgBlank
    Let pc5.Picture = ImgBlank
    Let pc6.Picture = ImgBlank
    Let pc1.Enabled = True
    Let pc2.Enabled = True
    Let pc3.Enabled = True
    Let pc4.Enabled = True
    Let pc5.Enabled = True
    Let pc6.Enabled = True
    Let TmrPop = True
    Let TmrTime1.Enabled = True
    Let cmdStart.Enabled = False
    Let lblover.Visible = False
End Sub

Private Sub Form_Load()
    Let pc1.Picture = imgBall
    Let pc2.Picture = imgBall
    Let pc3.Picture = imgBall
    Let pc4.Picture = imgBall
    Let pc5.Picture = imgBall
    Let pc6.Picture = imgBall
    Let Time = 120
    Let High = 0
    Call DisableCtrlAltDelete(True)
End Sub
Private Sub pc1_Click()
    If Popup = 1 And pc1.Picture = imgBall Then
        Let Win = Win + 1
        Let lblCount.Caption = Win
        Let pc1.Picture = imgPop
        Call Pop_Balloon
    End If
End Sub
Private Sub pc2_Click()
    If Popup = 2 And pc2.Picture = imgBall Then
        Let Win = Win + 1
        Let lblCount.Caption = Win
        Let pc2.Picture = imgPop
        Call Pop_Balloon
    End If
End Sub
Private Sub pc3_Click()
    If Popup = 3 And pc3.Picture = imgBall Then
        Let Win = Win + 1
        Let lblCount.Caption = Win
        Let pc3.Picture = imgPop
        Call Pop_Balloon
    End If
End Sub
Private Sub pc4_Click()
    If Popup = 4 And pc4.Picture = imgBall Then
        Let Win = Win + 1
        Let lblCount.Caption = Win
        Let pc4.Picture = imgPop
        Call Pop_Balloon
    End If
End Sub
Private Sub pc5_Click()
    If Popup = 5 And pc5.Picture = imgBall Then
        Let Win = Win + 1
        Let lblCount.Caption = Win
        Let pc5.Picture = imgPop
        Call Pop_Balloon
    End If
End Sub
Private Sub pc6_Click()
    If Popup = 6 And pc6.Picture = imgBall Then
        Let Win = Win + 1
        Let lblCount.Caption = Win
        Let pc6.Picture = imgPop
        Call Pop_Balloon
    End If
End Sub

Private Sub TmrPop_Timer()
    Popup = Int(Rnd * 6) + 1
    If Popup = 1 Then
        Let pc1.Picture = imgBall
        Let pc2.Picture = ImgBlank
        Let pc3.Picture = ImgBlank
        Let pc4.Picture = ImgBlank
        Let pc5.Picture = ImgBlank
        Let pc6.Picture = ImgBlank
    ElseIf Popup = 2 Then
        Let pc2.Picture = imgBall
        Let pc1.Picture = ImgBlank
        Let pc3.Picture = ImgBlank
        Let pc4.Picture = ImgBlank
        Let pc5.Picture = ImgBlank
        Let pc6.Picture = ImgBlank
    ElseIf Popup = 3 Then
        Let pc3.Picture = imgBall
        Let pc1.Picture = ImgBlank
        Let pc2.Picture = ImgBlank
        Let pc4.Picture = ImgBlank
        Let pc5.Picture = ImgBlank
        Let pc6.Picture = ImgBlank
    ElseIf Popup = 4 Then
        Let pc4.Picture = imgBall
        Let pc2.Picture = ImgBlank
        Let pc3.Picture = ImgBlank
        Let pc1.Picture = ImgBlank
        Let pc5.Picture = ImgBlank
        Let pc6.Picture = ImgBlank
    ElseIf Popup = 5 Then
        Let pc2.Picture = ImgBlank
        Let pc3.Picture = ImgBlank
        Let pc4.Picture = ImgBlank
        Let pc1.Picture = ImgBlank
        Let pc6.Picture = ImgBlank
        Let pc5.Picture = imgBall
    ElseIf Popup = 6 Then
        Let pc6.Picture = imgBall
        Let pc2.Picture = ImgBlank
        Let pc3.Picture = ImgBlank
        Let pc4.Picture = ImgBlank
        Let pc5.Picture = ImgBlank
        Let pc1.Picture = ImgBlank
    End If
    Let TmrPop.Enabled = False
    Let tmrpop2.Enabled = True
    
End Sub

Private Sub tmrpop2_Timer()
    Popup = Int(Rnd * 6) + 1
    If Popup = 1 Then
        Let pc1.Picture = imgBall
        Let pc2.Picture = ImgBlank
        Let pc3.Picture = ImgBlank
        Let pc4.Picture = ImgBlank
        Let pc5.Picture = ImgBlank
        Let pc6.Picture = ImgBlank
    ElseIf Popup = 2 Then
        Let pc2.Picture = imgBall
        Let pc1.Picture = ImgBlank
        Let pc3.Picture = ImgBlank
        Let pc4.Picture = ImgBlank
        Let pc5.Picture = ImgBlank
        Let pc6.Picture = ImgBlank
    ElseIf Popup = 3 Then
        Let pc3.Picture = imgBall
        Let pc2.Picture = ImgBlank
        Let pc1.Picture = ImgBlank
        Let pc4.Picture = ImgBlank
        Let pc5.Picture = ImgBlank
        Let pc6.Picture = ImgBlank
    ElseIf Popup = 4 Then
        Let pc4.Picture = imgBall
        Let pc2.Picture = ImgBlank
        Let pc3.Picture = ImgBlank
        Let pc1.Picture = ImgBlank
        Let pc5.Picture = ImgBlank
        Let pc6.Picture = ImgBlank
    ElseIf Popup = 5 Then
        Let pc5.Picture = imgBall
        Let pc2.Picture = ImgBlank
        Let pc3.Picture = ImgBlank
        Let pc4.Picture = ImgBlank
        Let pc1.Picture = ImgBlank
        Let pc6.Picture = ImgBlank
    ElseIf Popup = 6 Then
        Let pc6.Picture = imgBall
        Let pc2.Picture = ImgBlank
        Let pc3.Picture = ImgBlank
        Let pc4.Picture = ImgBlank
        Let pc5.Picture = ImgBlank
        Let pc1.Picture = ImgBlank
    End If
    Let TmrPop.Enabled = True
    Let tmrpop2.Enabled = False
End Sub

Private Sub TmrTime1_Timer()
    Let Time = Time - 1
    Let lblTimeLeft.Caption = Time
    If Time <= 0 Then
        pc1.Enabled = False
        pc2.Enabled = False
        pc3.Enabled = False
        pc4.Enabled = False
        pc5.Enabled = False
        pc6.Enabled = False
        Let TmrTime1.Enabled = False
        Let TmrTime2.Enabled = False
        Let TmrPop.Enabled = False
        Let tmrpop2.Enabled = False
        Let Time = 120
        Let cmdStart.Enabled = True
        Let cmdStart.Caption = "START NEW GAME"
        Let lblCount.Caption = 0
        If Win > High Then
            Let High = Win
            Let lblFinish.Caption = "Highest Score " & Win
        End If
        Let lblTimeLeft.Caption = 0
        Let High = Win
        Let Win = 0
        Let lblover.Visible = True
   End If
End Sub
Private Sub TmrTime2_Timer()
    Let Time = Time - 1
    Let lblTimeLeft.Caption = Time
    If Time <= 0 Then
        pc1.Enabled = False
        pc2.Enabled = False
        pc3.Enabled = False
        pc4.Enabled = False
        pc5.Enabled = False
        pc6.Enabled = False
        Let TmrTime1.Enabled = False
        Let TmrTime2.Enabled = False
        Let TmrPop.Enabled = False
        Let tmrpop2.Enabled = False
        Let Time = 120
        Let cmdStart.Enabled = True
        Let cmdStart.Caption = "START NEW GAME"
        Let lblCount.Caption = 0
        If Win > High Then
            Let High = Win
            Let lblFinish.Caption = "Highest Score " & Win
        End If
        Let lblTimeLeft.Caption = 0
        Let Win = 0
        Let lblover.Visible = True
    End If
End Sub


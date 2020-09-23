VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrLoad3 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   840
      Top             =   3480
   End
   Begin VB.Timer tmrLoad2 
      Enabled         =   0   'False
      Interval        =   750
      Left            =   120
      Top             =   3480
   End
   Begin VB.Timer tmrLoad1 
      Enabled         =   0   'False
      Interval        =   1350
      Left            =   480
      Top             =   3480
   End
   Begin VB.Label lblMess 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      Caption         =   "BALLOON POP IS LOADING PLEASE WAIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.Image imgBall 
      Height          =   2520
      Left            =   2760
      Picture         =   "frmSplash.frx":0000
      Stretch         =   -1  'True
      Top             =   1200
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Image imgPop 
      Height          =   2520
      Left            =   2760
      Picture         =   "frmSplash.frx":0655
      Stretch         =   -1  'True
      Top             =   1200
      Visible         =   0   'False
      Width           =   1635
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Let tmrLoad1.Enabled = True
    
End Sub

Private Sub tmrLoad1_Timer()
    Let imgBall.Visible = True
    Let tmrLoad1.Enabled = False
    Let tmrLoad2.Enabled = True
End Sub

Private Sub tmrLoad2_Timer()
    Let lblMess.Visible = True
    Let imgPop.Visible = True
    Let imgBall.Visible = False
    Let tmrLoad2.Enabled = False
    Let tmrLoad3.Enabled = True
End Sub

Private Sub tmrLoad3_Timer()
    Let tmrLoad3.Enabled = False
    frmBalloonPop.Show
    Unload frmSplash
End Sub

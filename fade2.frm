VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Circle Gradient"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   289
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   306
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   3480
      Width           =   255
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   4320
      ScaleHeight     =   795
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   3480
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808000&
      Height          =   855
      Left            =   3960
      ScaleHeight     =   795
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   3480
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3375
      Left            =   0
      MousePointer    =   2  'Cross
      ScaleHeight     =   221
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   301
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "Click anywhere on the picture to view to circle gradient."
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Click to change Color"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   3600
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************'
'*          This code was made by Nick Smith         *'
'*                  Copyright 2000                   *'
'*                                                   *'
'*        Questions or comments?  Send mail to       *'
'*               CCSkater@mailcity.com               *'
'*****************************************************'

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
Draw_GradientCircle Picture1, Picture1.ScaleWidth / 2, Picture1.ScaleHeight / 2, Picture2.BackColor, Picture3.BackColor
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
X_1 = x
Y_1 = y
Picture1.Cls
Draw_GradientCircle Picture1, X_1, Y_1, Picture2.BackColor, Picture3.BackColor
End Sub

Private Sub Picture2_Click()
CommonDialog1.ShowColor
Picture2.BackColor = CommonDialog1.Color
End Sub

Private Sub Picture3_Click()
CommonDialog1.ShowColor
Picture3.BackColor = CommonDialog1.Color
End Sub

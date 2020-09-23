VERSION 5.00
Begin VB.Form FrmAbout1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About......"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   3645
      Left            =   30
      Picture         =   "FrmAbout.frx":0000
      ScaleHeight     =   3585
      ScaleWidth      =   1005
      TabIndex        =   4
      Top             =   15
      Width           =   1065
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   435
      Left            =   3450
      TabIndex        =   1
      Top             =   3075
      Width           =   1110
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000003&
      X1              =   1110
      X2              =   1110
      Y1              =   45
      Y2              =   2355
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   4605
      X2              =   1110
      Y1              =   45
      Y2              =   45
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00E0E0E0&
      Height          =   2310
      Left            =   1110
      Top             =   45
      Width           =   3525
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3975
      Picture         =   "FrmAbout.frx":29DC
      Top             =   195
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Writen By Ben Jones"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1680
      TabIndex        =   3
      Top             =   1575
      Width           =   1485
   End
   Begin VB.Label Label2 
      Caption         =   "Windows 95 and Windows 98 Basic Programming Language"
      Height          =   540
      Left            =   1245
      TabIndex        =   2
      Top             =   840
      Width           =   3045
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Dreams DM ++"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1230
      TabIndex        =   0
      Top             =   495
      Width           =   1530
   End
End
Attribute VB_Name = "FrmAbout1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    Unload FrmAbout1
    Form1.Show
    
End Sub

Private Sub Form_Load()
    FrmAbout1.Icon = Form1.Icon
    
End Sub


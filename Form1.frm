VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5085
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "+"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "-"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin Proyecto1.uAOProgress uAOProgress 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   3615
      _ExtentX        =   2143
      _ExtentY        =   661
      Max             =   200
      MinDanger       =   50
      Value           =   100
      BackgroundColor =   4210752
      BackgroundDangerColor=   16777215
      BackColor       =   192
      BackAddColor    =   49152
      BackDangerColor =   8421631
      BackSubColor    =   128
      BorderColor     =   16711935
      ShowText        =   0   'False
      ShowShadow      =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command 
      Caption         =   "rnd"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ProgressBar_GotFocus()

End Sub

Private Sub Command_Click()
    uAOProgress.Value = Val(Rnd(1) * 100)
End Sub

Private Sub Command1_Click()
    uAOProgress.Value = uAOProgress.Value / 2
End Sub

Private Sub Command2_Click()
    uAOProgress.Value = uAOProgress.Value + (uAOProgress.Value / 2)
End Sub


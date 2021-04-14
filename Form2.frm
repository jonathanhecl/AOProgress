VERSION 5.00
Object = "{33D38DA7-F4D2-4EDB-85C4-4DC9E7E096EB}#5.0#0"; "AOProgress.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   800
      Left            =   6840
      Top             =   1320
   End
   Begin AOProgress.uAOProgress uAOProgress1 
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   1085
      Min             =   1
      MinDanger       =   50
      Value           =   1
      BackgroundDangerColor=   255
      BackColor       =   255
      BackAddColor    =   32768
      BackSubColor    =   128
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOProgress.uAOProgress uAOProgress2 
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   1085
      Min             =   1
      MinDanger       =   80
      Value           =   90
      BackgroundDangerColor=   12632256
      BackColor       =   16711680
      BackAddColor    =   16384
      BackSubColor    =   4194368
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOProgress.uAOProgress uAOProgress3 
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   1085
      Min             =   1
      Value           =   1
      BackColor       =   65280
      BackAddColor    =   32768
      BackSubColor    =   128
      CustomText      =   "Cargando..."
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOProgress.uAOProgress uAOProgress4 
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      Min             =   1
      Value           =   1
      Animate         =   0   'False
      BackColor       =   33023
      BackAddColor    =   4934475
      BackSubColor    =   8224125
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOProgress.uAOProgress uAOProgress5 
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   3240
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   1085
      Min             =   1
      Value           =   1
      UseBackground   =   0   'False
      BackColor       =   16711935
      BackAddColor    =   4934475
      BackSubColor    =   8224125
      ShowText        =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    uAOProgress1.Value = Val(Rnd(1) * 100)
    uAOProgress2.Value = Val(Rnd(1) * 100)
    uAOProgress3.Value = Val(Rnd(1) * 100)
    uAOProgress4.Value = Val(Rnd(1) * 100)
    uAOProgress5.Value = Val(Rnd(1) * 100)
End Sub



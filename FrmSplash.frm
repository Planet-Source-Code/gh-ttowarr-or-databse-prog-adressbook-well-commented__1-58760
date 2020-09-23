VERSION 5.00
Begin VB.Form FrmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Tmr 
      Interval        =   3000
      Left            =   120
      Top             =   2400
   End
   Begin VB.Label LblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Biblio Pro 1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label LblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "All copyrights to Gh3ttoWarr!or"
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   2
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Label LblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "HardStream Productions"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label LblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Made by Gh3ttoWarr!or"
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Shape ShpBox 
      Height          =   3015
      Left            =   0
      Top             =   0
      Width           =   4095
   End
   Begin VB.Image ImgUsers 
      Height          =   1440
      Left            =   120
      Picture         =   "FrmSplash.frx":0000
      Top             =   120
      Width           =   1440
   End
   Begin VB.Image ImgHsd 
      Height          =   1290
      Left            =   2880
      Picture         =   "FrmSplash.frx":70CA
      Top             =   120
      Width           =   1065
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Tmr_Timer()
Unload Me
FrmMain.Show
End Sub

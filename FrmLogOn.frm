VERSION 5.00
Begin VB.Form FrmLogOn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please login"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   3495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Txtpass 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Line Line 
      X1              =   0
      X2              =   3480
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Shape ShpBox 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   1320
      Top             =   0
      Width           =   2175
   End
   Begin VB.Image ImgKeys 
      Height          =   1440
      Left            =   0
      Picture         =   "FrmLogOn.frx":0000
      Top             =   0
      Width           =   1440
   End
End
Attribute VB_Name = "FrmLogOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancel_Click()
End
End Sub

Private Sub CmdOk_Click()
If Txtpass.Text = "pass" Then
FrmMain.Show
Unload Me
End If
End Sub

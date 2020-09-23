VERSION 5.00
Begin VB.Form FrmNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New User"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3480
   Icon            =   "FrmNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   3480
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Frame Frame 
      Caption         =   "New User"
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   3255
      Begin VB.TextBox Txt 
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   4200
         Width           =   3015
      End
      Begin VB.TextBox Txt 
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   3480
         Width           =   3015
      End
      Begin VB.TextBox Txt 
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   2760
         Width           =   3015
      End
      Begin VB.TextBox Txt 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Width           =   3015
      End
      Begin VB.TextBox Txt 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox Txt 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label LblInfo 
         Caption         =   "Phone:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   14
         Top             =   3960
         Width           =   3015
      End
      Begin VB.Label LblInfo 
         Caption         =   "Zip:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   3240
         Width           =   3015
      End
      Begin VB.Label LblInfo 
         Caption         =   "State:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   3015
      End
      Begin VB.Label LblInfo 
         Caption         =   "City:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   3015
      End
      Begin VB.Label LblInfo 
         Caption         =   "Adress:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label LblInfo 
         Caption         =   "Name:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   3015
      End
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
   Begin VB.Line Line 
      X1              =   0
      X2              =   3480
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Image ImgEdit 
      Height          =   1440
      Left            =   0
      Picture         =   "FrmNew.frx":0ECA
      Top             =   0
      Width           =   1440
   End
End
Attribute VB_Name = "FrmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==========================================================
'CmdCancel_Click (Cancels edit of user)
'==========================================================

Private Sub CmdCancel_Click()
Unload Me
End Sub

'==========================================================
'CmdOk_Click (Submit the new user)
'==========================================================

Private Sub CmdOk_Click()
Dim i As Integer

For i = 0 To 5

    If Txt(i).Text = "" Then
        MsgBox "Please fill in all textfields...", vbInformation, "Incomplete..."
        Exit Sub
    End If
Next i

FrmMain.AddUser Txt(0), Txt(1), Txt(2), Txt(3), Txt(4), Txt(5)
FrmMain.RefreshList
Unload Me
End Sub

'==========================================================
'Form unLoad (When this form unloads)
'==========================================================

Private Sub Form_Unload(Cancel As Integer)
FrmMain.RefreshList
End Sub

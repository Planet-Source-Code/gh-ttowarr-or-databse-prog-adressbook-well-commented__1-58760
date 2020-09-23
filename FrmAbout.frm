VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Biblio 1.0"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3480
   Icon            =   "FrmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   3480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Caption         =   "About"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3255
      Begin VB.Label LblInfo 
         Caption         =   "Made by Gh3ttoWarr!or,           HardStream Productions"
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   3015
      End
   End
   Begin VB.Line Line 
      X1              =   0
      X2              =   3480
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Image ImgEdit 
      Height          =   720
      Left            =   120
      Picture         =   "FrmAbout.frx":0ECA
      Top             =   120
      Width           =   720
   End
   Begin VB.Shape ShpBox 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


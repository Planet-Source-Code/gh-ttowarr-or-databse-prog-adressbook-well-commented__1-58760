VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Biblio Pro 1.0"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   8280
   Icon            =   "FrmMain.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Caption         =   "User Options"
      Height          =   855
      Index           =   2
      Left            =   3240
      TabIndex        =   3
      Top             =   6240
      Width           =   4935
      Begin VB.CommandButton CmdEdit 
         Caption         =   "Edit"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Print"
         Height          =   495
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdRemove 
         Caption         =   "Remove"
         Height          =   495
         Left            =   2520
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "New"
         Height          =   495
         Left            =   3720
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "List of Users"
      Height          =   5535
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   3015
      Begin VB.ListBox List 
         Height          =   5130
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Current User"
      Height          =   4575
      Index           =   0
      Left            =   3240
      TabIndex        =   0
      Top             =   1560
      Width           =   4935
      Begin VB.Label LblDbInfo 
         Caption         =   "."
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   19
         Top             =   2160
         Width           =   3615
      End
      Begin VB.Label LblDbInfo 
         Caption         =   "."
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   18
         Top             =   1800
         Width           =   3615
      End
      Begin VB.Label LblDbInfo 
         Caption         =   "."
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   17
         Top             =   1440
         Width           =   3615
      End
      Begin VB.Label LblDbInfo 
         Caption         =   "."
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   16
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Label LblDbInfo 
         Caption         =   "."
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   15
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label LblDbInfo 
         Caption         =   "."
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   14
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label LblInfo 
         Caption         =   "Phone:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label LblInfo 
         Caption         =   "Zip:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label LblInfo 
         Caption         =   "State:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label LblInfo 
         Caption         =   "City:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label LblInfo 
         Caption         =   "Adress:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   615
      End
      Begin VB.Label LblInfo 
         Caption         =   "Name:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Image ImgLogo 
      Height          =   1290
      Left            =   6960
      Picture         =   "FrmMain.frx":0ECA
      Top             =   120
      Width           =   1065
   End
   Begin VB.Image ImgUsers 
      Height          =   1440
      Left            =   0
      Picture         =   "FrmMain.frx":579C
      Top             =   0
      Width           =   1440
   End
   Begin VB.Line Line 
      X1              =   0
      X2              =   8280
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Shape ShpBox 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   0
      Top             =   0
      Width           =   8295
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuQuit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu MnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public db As Database
Public rstInfo As Recordset

'==========================================================
'CmdEdit_Click (If the user clicks on the button "Edit")
'==========================================================

Private Sub CmdEdit_Click()
If List.ListIndex = -1 Then
MsgBox "Please select first an user you want to edit..", vbInformation, "Nothing selected.."
Exit Sub
End If

FrmEdit.Show
End Sub

'==========================================================
'CmdNew_Click (If the user clicks on the button "New")
'==========================================================

Private Sub CmdNew_Click()
FrmNew.Show
End Sub

'==========================================================
'CmdPrint_Click (If the user clicks on the button "Print")
'==========================================================

Private Sub CmdPrint_Click()
If List.ListIndex = -1 Then
MsgBox "Please select first an user you want to print..", vbInformation, "Nothing selected.."
Exit Sub
End If

PrintUser
End Sub

'==========================================================
'CmdRemove_Click (If the user clicks on the button "Remove")
'==========================================================

Private Sub CmdRemove_Click()
Dim i As Integer
If List.ListIndex = -1 Then
MsgBox "Please select first an user you want to remove..", vbInformation, "Nothing selected.."
Exit Sub
End If

rstInfo.Delete
RefreshList

For i = 0 To LblDbInfo.Count - 1
    LblDbInfo(i).Caption = "."
Next i
End Sub

'==========================================================
'Form Load (When this form loads)
'==========================================================

Private Sub Form_Load()
Set db = OpenDatabase(App.Path & "\Other\Database.mdb")
Set rstInfo = db.OpenRecordset("Info")
RefreshList
End Sub

'==========================================================
'Add User (started by FrmNew)
'==========================================================

Public Sub AddUser(Name As String, Adress As String, City As String, State As String, Zip As String, Phone As String)
With rstInfo

.AddNew
    !Name = Name
    !Address = Adress
    !City = City
    !State = State
    !Zip = Zip
    !Phone = Phone
.Update

End With
End Sub

'==========================================================
'Edit User (started by FrmEdit)
'==========================================================

Public Sub EditUser(Name As String, Adress As String, City As String, State As String, Zip As String, Phone As String)
With rstInfo

.Edit
    !Name = Name
    !Address = Adress
    !City = City
    !State = State
    !Zip = Zip
    !Phone = Phone
.Update

End With
End Sub

'==========================================================
'Refresh List (The list with all the users)
'==========================================================

Public Sub RefreshList()
Dim i As Integer
List.Clear

With rstInfo

If .RecordCount = 0 Then Exit Sub

.MoveFirst

For i = 0 To .RecordCount
If .EOF Then Exit Sub
List.AddItem !Name
.MoveNext
Next i

End With
End Sub

'==========================================================
'List Click (load new user info)
'==========================================================

Private Sub List_Click()
With rstInfo
.MoveFirst
.Move List.ListIndex

.Edit
    LblDbInfo(0).Caption = !Name
    LblDbInfo(1).Caption = !Address
    LblDbInfo(2).Caption = !City
    LblDbInfo(3).Caption = !State
    LblDbInfo(4).Caption = !Zip
    LblDbInfo(5).Caption = !Phone
.Update

End With

End Sub

'==========================================================
'Print User (print the current user info)
'==========================================================

Public Sub PrintUser()
Dim i As Integer
For i = 0 To 5
Printer.Print (LblDbInfo(i).Caption)
Next i
Printer.EndDoc
End Sub

'==========================================================
'Menu Click About(show about)
'==========================================================

Private Sub MnuAbout_Click()
FrmAbout.Show
End Sub

'==========================================================
'Menu Click Quit(Quit)
'==========================================================

Private Sub MnuQuit_Click()
End
End Sub

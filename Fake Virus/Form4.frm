VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please wait...."
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4950
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1320
      Top             =   840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Please wait while window Deleting C:\ Drive Contents."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   -360
      TabIndex        =   0
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Dim hSysMenu As Long
hSysMenu = GetSystemMenu(hwnd, False)
RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMMAND
End Sub

Private Sub Timer1_Timer()
ProgressBar1 = ProgressBar1 + 1
Label1 = "Window preparing to install 007 Trojan Virus" & vbCrLf & ProgressBar1 & "%"
If ProgressBar1 = 100 Then
Form5.Show
Unload Form4
End If

End Sub

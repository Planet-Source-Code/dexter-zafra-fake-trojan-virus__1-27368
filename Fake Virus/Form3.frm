VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window warning!"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3000
      TabIndex        =   2
      Top             =   1140
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&No"
      Height          =   400
      Left            =   1740
      TabIndex        =   1
      Top             =   1140
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Yes"
      Height          =   400
      Left            =   480
      TabIndex        =   0
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Are you sure you want to delete all folders and files from this computer?"
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Sub Sound()
sndPlaySound App.Path & "\" & "ret.wav", &H1
End Sub


Private Sub Command1_Click()
   Unload Me
   Form1.Show
   remItem Form1
   
End Sub

Private Sub Command2_Click()
   Unload Me
   Form1.Show
   remItem Form1
   
End Sub

Private Sub Command3_Click()
   Unload Me
   Form1.Show
   remItem Form1
   
End Sub

Private Sub Form_Load()
Dim hSysMenu As Long
hSysMenu = GetSystemMenu(hwnd, False)
RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMMAND

Sound
   
End Sub



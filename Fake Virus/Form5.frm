VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window prompt.."
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4380
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   4380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&No"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Yes"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "007 Trojan Virus has successfully installed. Would you like to locate the 007 Trojan file and Delete it?"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Sub Sound()
sndPlaySound App.Path & "\" & "ret.wav", &H1
End Sub

Private Sub Command1_Click()
Form8.Show
Unload Me
End Sub

Private Sub Command2_Click()
Form8.Show
Unload Me
End Sub

Private Sub Form_Load()
Dim hSysMenu As Long
hSysMenu = GetSystemMenu(hwnd, False)
RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMMAND
Sound
End Sub

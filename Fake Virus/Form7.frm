VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window Critical Information!"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Window could not start without the deleted files.Would you like to reinstall all the deleted files back to your C:\Drive?"""
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Sub Sound()
sndPlaySound App.Path & "\" & "ret.wav", &H1
End Sub


Private Sub Command1_Click()
Form4.Show
Unload Form7
End Sub

Private Sub Form_Load()
Dim hSysMenu As Long
hSysMenu = GetSystemMenu(hwnd, False)
RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMMAND

Sound

End Sub


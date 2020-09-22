VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window Critical information!"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Window  recommend you to quit running this program."
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Window has detected a malicious file while reinstalling the C:\ Drive contents."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Sub Sound()
sndPlaySound App.Path & "\" & "ret.wav", &H1
End Sub

Private Sub Command1_Click()
Formscroll.Show
Unload Me
End Sub

Private Sub Form_Load()
Dim hSysMenu As Long
hSysMenu = GetSystemMenu(hwnd, False)
RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMMAND

Sound
End Sub

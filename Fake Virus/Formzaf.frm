VERSION 5.00
Begin VB.Form Formzaf 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   6885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11505
   LinkTopic       =   "Form3"
   ScaleHeight     =   6885
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "This is just a joke..So pls! don't get mad if my program annoyed you..Thank you..Have a great day..Byeeee till next time.."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   2520
      Width           =   7695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "To close this program.Just click anywhere in the form..Thank you!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   6360
      Width           =   6375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RETXEDZ007"
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   10215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This is not a real Trojan Virus..So you don't have to worry..It won't harm your computer..I Love you (KM..DZ)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   3120
      TabIndex        =   0
      Top             =   3120
      Width           =   6015
   End
End
Attribute VB_Name = "Formzaf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Click()
Form6.Show
End Sub

Private Sub Label1_Click()
Form6.Show
End Sub
Private Sub Label2_Click()
Form6.Show
End Sub

Private Sub Label3_Click()
Form6.Show
End Sub

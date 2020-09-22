VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   9750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11730
   LinkTopic       =   "Form2"
   ScaleHeight     =   9750
   ScaleWidth      =   11730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Left            =   1320
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Left            =   780
      Top             =   120
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Window could not run File not found 'Run.DLL'"
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
      Left            =   2160
      TabIndex        =   3
      Top             =   4440
      Width           =   8775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "System Malfuntion Window could not run with out the C:\Drive files."
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
      Height          =   735
      Left            =   1800
      TabIndex        =   2
      Top             =   3120
      Width           =   9375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait while window is updating the C:\Drive of your computer."
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
      Height          =   735
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Width           =   8175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "All of the C:\Drive files have been successfully  deleted. "
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
      Left            =   2040
      TabIndex        =   0
      Top             =   1800
      Width           =   9195
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Label1.Visible = False
   Label2.Visible = False
   Label4.Visible = False
   
   Timer2.Enabled = True
   Timer2.Interval = 8000
                                                        
   Timer3.Enabled = True
   Timer3.Interval = 13000
                            
End Sub

Private Sub Timer2_Timer()
   Label1.Visible = True
   Label2.Visible = True
   Label4.Visible = True
End Sub

Private Sub Timer3_Timer()
Form7.Show
Unload Me
End Sub

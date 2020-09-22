VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Preparing to Delete Files"
   ClientHeight    =   6915
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6915
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lstFile 
      Height          =   6075
      Left            =   3720
      TabIndex        =   7
      Top             =   480
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   10716
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command5 
      Height          =   315
      Left            =   4620
      Picture         =   "Form1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   60
      Width           =   315
   End
   Begin VB.CommandButton Command4 
      Height          =   315
      Left            =   4260
      Picture         =   "Form1.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   60
      Width           =   315
   End
   Begin VB.CommandButton Command3 
      Height          =   315
      Left            =   3900
      Picture         =   "Form1.frx":0646
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   60
      Width           =   315
   End
   Begin VB.CommandButton Command2 
      Height          =   315
      Left            =   3120
      Picture         =   "Form1.frx":0748
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   60
      Width           =   315
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   2955
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   6600
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "41 object(s)"
            TextSave        =   "41 object(s)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17886
            Text            =   "15MB (Disk Free Space 658MB)"
            TextSave        =   "15MB (Disk Free Space 658MB)"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView treFile 
      Height          =   6135
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   10821
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   60
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":084A
            Key             =   "Desktop"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0C9C
            Key             =   "MyComp"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10EE
            Key             =   "drive"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1540
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1652
            Key             =   "open"
         EndProperty
      EndProperty
   End
   Begin VB.Line Line2 
      DrawMode        =   4  'Mask Not Pen
      X1              =   3060
      X2              =   13620
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Line Line1 
      DrawMode        =   4  'Mask Not Pen
      X1              =   60
      X2              =   13620
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
    Private Const SPI_SCREENSAVERRUNNING = 97

Private Sub Form_Load()
   Dim ret As Integer
    Dim pOld As Boolean
    ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
   Dim fs As FileSystemObject, dc As Drives, d As Drive, s As String
   Dim sf As Folder, i As Long, f As Folders, f1 As Folder
   Dim fil As File, fi As Files, fi1 As File
   Set fs = CreateObject("scripting.filesystemobject")
   Set dc = fs.Drives
      
   treFile.Nodes.Add , tvwFirst, "Desktop", "Desktop", "Desktop"
   treFile.Nodes.Add , tvwNext, "MyComp", "My Computer", "MyComp"
   treFile.Nodes.Add "MyComp", tvwChild, "(A:)", "(A:)", "drive"
   For Each d In dc
      If d.DriveType <> Remote And d.DriveLetter <> "A" Then
         s = "(" & d.Path & ")"
         On Error Resume Next
         treFile.Nodes.Add "MyComp", tvwChild, s, s, "drive"
         Set sf = fs.GetFolder(d.RootFolder)
         Set f = sf.SubFolders
         Set fil = fs.GetFile(d.RootFolder)
         For Each f1 In f
            treFile.Nodes.Add s, tvwChild, f1.Name, f1.Name, "closed"
            If d.DriveLetter = "C" Then
              lstFile.ListItems.Add , f1.Name, f1.Name, , "closed"
            End If
         Next
      End If
   Next
      
   Me.Show
   treFile.Nodes.Item("MyComp").Expanded = True
   treFile.Nodes.Item("(C:)").Expanded = True
   Form3.Show vbModal
     
End Sub


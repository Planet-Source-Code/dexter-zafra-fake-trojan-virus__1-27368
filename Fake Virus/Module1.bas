Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function ExitWindowsEx Lib "user32" _
            (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
    Public Const EWX_FORCE = 4
    Public Const EWX_REBOOT = 2
    Public Const EWX_SHUTDOWN = 1


Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" _
            (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, _
             ByVal bFailIfExists As Long) As Long

Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long


Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    Public Const SC_CLOSE = &HF060&
    Public Const MF_BYCOMMAND = &H0&

' ***FSO = File system object***

Public Function remItem(frm As Form)
   Dim fs As FileSystemObject, dc As Drives, d As Drive, s As String
   Dim sf As Folder, f As Folders, f1 As Folder, i As Long
   
   ' Set the FSO and get the get the drives from it.
   ' It is the actual file structure of the computer the app resides on.
   Set fs = CreateObject("scripting.filesystemobject")
   Set dc = fs.Drives
         
'***Use the form passed for the following statements***
   With frm
      For Each d In dc   ' Get the properties for each drive in the FSO
         ' Only get local drives.  If you try to get all drives it may take a long time
         ' to load depending on how many remote drives there are and how many folders are
         ' within the remote.
         If d.DriveType <> Remote And d.DriveLetter <> "A" Then
            s = "(" & d.Path & ")"   ' Gets the path of the current drive.  Ex: "C:\"
            Set sf = fs.GetFolder(d.RootFolder) ' Gets all the sub folders within the drive/directory.
            Set f = sf.SubFolders ' Sets the subfolders to a variable.
            On Error Resume Next
            For Each f1 In f   ' Looping through all the subfolders within the drive/directory
               !lstFile.ListItems.Add , f1.Name, f1.Name, , "closed"  ' Adds the Subfolder to the list view with the closed folder as an icon.
            Next
            For Each f1 In f ' Looping through all the subfolders again.  It seems redundant, but it is necessary to make it appear as though
                             ' the folders are being deleted from the computer.
               !treFile.Nodes.Item("MyComp").Expanded = True ' Make sure the MyComp Node is expanded
               !treFile.Nodes.Item(s).Expanded = True  ' Makes sure the current drive Node is expanded.
               i = 0
               Do Until i = 10 ' The loop just slows down the process to make the appearance of the folders being deleted visible to the user.
                  !treFile.Nodes.Remove f1.Name   ' Remove the Node from the tree view.
                  !lstFile.ListItems.Remove f1.Name ' Remove the Item from the list view.
               i = i + 1
               .Refresh   ' The tree view refreshes itself, but the list view needs to be refreshed.  I've found it looks less choppy if the whole
                          ' form is refreshed.
               Loop
               
            Next
            !treFile.Nodes.Remove s  ' Remove the current drive node.
         End If
      Next
      !treFile.Nodes.Remove "(A:)"
      !treFile.Nodes.Remove "MyComp"
      !treFile.Nodes.Remove "Desktop"
   End With
   
   Formins.Show
   Unload Form1
End Function


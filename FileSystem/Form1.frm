VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FileSystem Demo"
   ClientHeight    =   7080
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   9264
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   9264
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1032
      Left            =   0
      ScaleHeight     =   984
      ScaleWidth      =   9216
      TabIndex        =   3
      Top             =   0
      Width           =   9264
      Begin VB.CommandButton cmdAbout 
         Caption         =   "About FileSystem"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   852
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   60
         Width           =   1212
      End
      Begin VB.CommandButton cmdReadMe 
         Caption         =   "Show ReadMe"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   852
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
         Width           =   1212
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   " Run FileSystem"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   852
         Left            =   2580
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   60
         Width           =   1212
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Exit DemoApp"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   852
         Left            =   7980
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Width           =   1212
      End
      Begin VB.CommandButton cmdList 
         Caption         =   "List Functions"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   852
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   60
         Width           =   1212
      End
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      ItemData        =   "Form1.frx":0000
      Left            =   5940
      List            =   "Form1.frx":001F
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1080
      Width           =   3252
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   5532
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1440
      Width           =   9072
   End
   Begin VB.Label Label2 
      Caption         =   "Pick a drive letter where you want the fileio functions to work"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   1110
      Width           =   5712
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdList_Click()
    ListFunctions
End Sub

Private Sub cmdReadMe_Click()
    ShowReadMe
End Sub

Private Sub cmdAbout_Click()
    ShowAbout
End Sub

Private Sub cmdRun_Click()
    If Combo1.Text = "" Then
        MsgBox "Please pick a drive letter first"
        On Error Resume Next
        Combo1.SetFocus
        Exit Sub
    End If
    RunThis Combo1.Text
End Sub

Private Sub RunThis(ByVal Drive As String)
    Dim Path As String
    Dim Filename As String
    Dim hFile As Long
    Dim Root As String
    Dim FilePath As String
    Dim TempPath As String
    Dim TempFile As String
    Dim Today As String
    Dim File As TFile
    Dim Start As Long
    
    Start = Timer
    Today = Now
    
    ' these are the names of files and folder we will use
    ' in our demo test
    Text1.Text = ""
    Root = Drive & "FileTest"
    Path = Root & "\Demo\Test"
    TempPath = GetPathName(Root) & "\Temporary"
    Filename = "Text.TXT"
    FilePath = Path & "\" & Filename
    TempFile = TempPath & "\" & Filename
    
    DebugOut "*** Starting tests...", 0
    DebugOut "Windows  : " & FileSystem.GetSpecialFolder(WindowsFolder)
    DebugOut "System   : " & FileSystem.GetSpecialFolder(SystemFolder)
    DebugOut "Temporary: " & FileSystem.GetSpecialFolder(TemporaryFolder)
    DebugOut "*** Calling some functions...", 0
    DebugOut "Testing GetTempFileName() function: " & FileSystem.GetTempFileName(, "rad")
    DebugOut "Testing LocalDateToUTC(Today) function: " & FileSystem.LocalDateToUTC(Today)
    DebugOut "Testing UTCToLocalDate(Today) function: " & FileSystem.UTCToLocalDate(FileSystem.LocalDateToUTC(Today))
    DebugOut vbCrLf
    DebugOut "*** Starting file handling functions...", 0
    DebugOut "FileSystem.CreateFolder(" & Path & ")", 2
    If FileSystem.CreateFolder(Path) Then
        DebugOut ">> Created folder " & Path
        DebugOut "CreateFolder(" & TempPath & ")", 2
        If CreateFolder(TempPath) Then
            DebugOut ">> Created folder " & TempPath
            DebugOut "CreateTextFile(" & FilePath & ")", 2
            hFile = CreateTextFile(FilePath)
            If hFile <> INVALID_HANDLE_VALUE Then
                DebugOut ">> Created text file " & FilePath
                DebugOut "CloseFile hFile", 2
                CloseFile hFile
                DebugOut "FileSystem.GetFile(" & FilePath & ", File)", 2
                If FileSystem.GetFile(FilePath, File) Then
                    DebugOut ">> File Information (" & FilePath & ")"
                    DebugOut ">> Name            : " & File.Name
                    DebugOut ">> Path            : " & File.Path
                    DebugOut ">> TypeName        : " & File.TypeName
                    DebugOut ">> Attributes      : " & File.Attributes
                    DebugOut ">> DateCreated     : " & File.DateCreated
                    DebugOut ">> DateLastAccessed: " & File.DateLastAccessed
                    DebugOut ">> DateLastModified: " & File.DateLastModified
                    DebugOut ">> Size            : " & File.Size
                    DebugOut ">> Alternate       : " & File.Alternate
                    DebugOut ">> Directory       : " & File.Directory
                    DebugOut ">> ParentFolder    : " & File.ParentFolder
                    DebugOut ">> DisplayName     : " & File.DisplayName
                    DebugOut ">> Extension       : " & File.Extension
                    DebugOut ">> ShortName       : " & File.ShortName
                    DebugOut ">> ShortPathName   : " & File.ShortPathName
                    DebugOut ">> FullPathName    : " & File.FullPathName
                    DebugOut ">> CompressedSize  : " & File.CompressedSize
                    DebugOut ">> Icon            : " & File.hIcon
                End If
                DebugOut "GetFileContentType(" & FilePath & ")", 2
                DebugOut ">> Content-type of " & FilePath & " is """ & GetFileContentType(FilePath) & Chr(34)
                DebugOut "FileSystem.MoveFile(" & Path & " & ""\*.*"", AddBkSlash(" & TempPath & "))", 2
                If FileSystem.MoveFile(Path & "\*.*", AddBkSlash(TempPath)) Then
                    DebugOut ">> Files moved from " & Path & " to " & TempPath
                    DebugOut "FileSystem.CopyFile(" & TempFile & ", " & FilePath & ", True)", 2
                    If FileSystem.CopyFile(TempFile, FilePath, True) Then
                        DebugOut ">> File " & TempFile & " copied to " & FilePath
                        DebugOut "FileSystem.DeleteFile(" & TempFile & ", True)", 2
                        If FileSystem.DeleteFile(TempFile, True) Then
                            DebugOut ">> File deleted " & TempFile
                        Else
                            DebugOut "!! Cannot delete file " & TempFile
                        End If
                    Else
                        DebugOut "!! Cannot copy file " & TempFile & " to " & FilePath
                    End If
                Else
                    DebugOut "!! Cannot move files from " & Path & " to " & TempPath
                End If
                DebugOut "FileSystem.MoveFolder(GetParentFolderName(" & Path & "), AddBkSlash(" & TempPath & "))", 2
                ' move demo folder to temporary folder
                If FileSystem.MoveFolder(GetParentFolderName(Path), AddBkSlash(TempPath)) Then
                    DebugOut ">> Folder " & GetParentFolderName(Path) & " was moved to " & TempPath
                Else
                    DebugOut "!! Cannot move folder " & GetParentFolderName(Path) & " to " & TempPath
                End If
                DebugOut "FileSystem.CopyFolder(" & TempPath & ", AddBkSlash(GetParentFolderName(" & Path & ")))", 2
                ' move demo folder to temporary folder
                If FileSystem.CopyFolder(TempPath, AddBkSlash(GetParentFolderName(Path))) Then
                    DebugOut ">> Folder " & TempPath & " was moved to " & GetParentFolderName(Path)
                Else
                    DebugOut "!! Cannot move folder " & TempPath & " to " & GetParentFolderName(Path)
                End If
                DebugOut "FileSystem.DeleteFolder(" & Root & ", True)", 2
                If FileSystem.DeleteFolder(Root, True) Then
                    DebugOut ">> Folder deleted " & Root
                Else
                    DebugOut "!! Cannot delete folder " & Root
                End If
                DebugOut "FileSystem.DeleteFolder(AddBkSlash(" & TempPath & ") & ""*.*"", True)", 2
                If FileSystem.DeleteFolder(AddBkSlash(TempPath) & "*.*", True) Then
                    DebugOut ">> Folders deleted in " & TempPath
                Else
                    DebugOut "!! Cannot delete folders in " & TempPath
                End If
                DebugOut "FileSystem.DeleteFolder(" & TempPath & ", True)", 2
                If FileSystem.DeleteFolder(TempPath, True) Then
                    DebugOut ">> Folder deleted " & TempPath
                Else
                    DebugOut "!! Cannot delete folder " & TempPath
                End If
            Else
                DebugOut "!! Cannot create " & FilePath
            End If
        Else
            DebugOut "!! Cannot create " & TempPath
        End If
    Else
        DebugOut "!! Cannot create " & Path
    End If
    DebugOut "*** End of tests", 0
    DebugOut "Elapsed time: " & (Timer - Start) & " secs."
    DebugOut vbCrLf
    DebugOut "--------------------------------------------------", 0
    DebugOut "Thank you for using filesystem module.", 0
    DebugOut "Report bug to marclei@spnorte.com", 0
    DebugOut "--------------------------------------------------", 0
End Sub

Private Sub ShowReadMe()
    Dim BufStr As String

    BufStr = BufStr & "FileSystem Module ReadMe" & vbCrLf
    BufStr = BufStr & "(february, 04 2001)" & vbCrLf
    BufStr = BufStr & "-------------------------------------------------------------------------" & vbCrLf
    BufStr = BufStr & "" & vbCrLf
    BufStr = BufStr & "This demo will create, move and delete folders and files in you hard disk" & vbCrLf
    BufStr = BufStr & "but don't be afraid. It will not harm your computer. The tests will " & vbCrLf
    BufStr = BufStr & "take place at <Drive you choose>\FileTest directory and " & vbCrLf
    BufStr = BufStr & "<Drive you choose>\Temporary directory. This way you can monitor the" & vbCrLf
    BufStr = BufStr & "actions when it is running. Just pick a drive letter and click the" & vbCrLf
    BufStr = BufStr & "<RUN> button. At the end of the demo all test files and folders will be." & vbCrLf
    BufStr = BufStr & "deleted. The best way to understand what is going on is debugging the " & vbCrLf
    BufStr = BufStr & "RunThis() method" & vbCrLf
    BufStr = BufStr & "" & vbCrLf
    BufStr = BufStr & "-------------------------------------------------------------------------" & vbCrLf
    BufStr = BufStr & "Contact the author" & vbCrLf
    BufStr = BufStr & "marclei@spnorte.com" & vbCrLf
    Text1.Text = BufStr
End Sub

Private Sub ListFunctions()
    Dim BufStr As String
    
    BufStr = BufStr & "FileSystem Module Function List" & vbCrLf
    BufStr = BufStr & "(february, 04 2001)" & vbCrLf
    BufStr = BufStr & "-------------------------------------------------------------------------" & vbCrLf
    BufStr = BufStr & "StripBkSlash()" & vbTab & vbTab
    BufStr = BufStr & "StripLtSlash()" & vbTab & vbTab
    BufStr = BufStr & "AddBkSlash()" & vbCrLf
    BufStr = BufStr & "AddLtSlash()" & vbTab & vbTab
    BufStr = BufStr & "SetCurrentDirectory()" & vbTab
    BufStr = BufStr & "GetParentFolderName()" & vbCrLf
    BufStr = BufStr & "GetFileAttributes()" & vbTab
    BufStr = BufStr & "SetFileAttributes()" & vbTab
    BufStr = BufStr & "FileExists()" & vbCrLf
    BufStr = BufStr & "FolderExists()" & vbTab & vbTab
    BufStr = BufStr & "GetFileName()" & vbTab & vbTab
    BufStr = BufStr & "GetBasename()" & vbCrLf
    BufStr = BufStr & "GetPathName()" & vbTab & vbTab
    BufStr = BufStr & "CreateFolder()" & vbTab & vbTab
    BufStr = BufStr & "DeleteFile()" & vbCrLf
    BufStr = BufStr & "CopyFile()" & vbTab & vbTab
    BufStr = BufStr & "MoveFile()" & vbTab & vbTab
    BufStr = BufStr & "CopyFolder()" & vbCrLf
    BufStr = BufStr & "MoveFolder()" & vbTab & vbTab
    BufStr = BufStr & "DeleteFolder()" & vbTab & vbTab
    BufStr = BufStr & "CreateTextFile()" & vbCrLf
    BufStr = BufStr & "FindFirst()" & vbTab & vbTab
    BufStr = BufStr & "FindNext()" & vbTab & vbTab
    BufStr = BufStr & "GetPictureInfo()" & vbCrLf
    BufStr = BufStr & "GetExtensionName()" & vbTab
    BufStr = BufStr & "OpenTextFile()" & vbTab & vbTab
    BufStr = BufStr & "CloseFile()" & vbCrLf
    BufStr = BufStr & "GetFileContentType()" & vbTab
    BufStr = BufStr & "BuildPath()" & vbTab & vbTab
    BufStr = BufStr & "GetFullPathName()" & vbCrLf
    BufStr = BufStr & "IsReadOnly()" & vbTab & vbTab
    BufStr = BufStr & "IsHidden()" & vbTab & vbTab
    BufStr = BufStr & "IsArchive()" & vbCrLf
    BufStr = BufStr & "IsSystem()" & vbTab & vbTab
    BufStr = BufStr & "IsTemporary()" & vbTab & vbTab
    BufStr = BufStr & "IsCompressed()" & vbCrLf
    BufStr = BufStr & "IsDirectory()" & vbTab & vbTab
    BufStr = BufStr & "GetShortPathName()" & vbTab
    BufStr = BufStr & "GetTempPath()" & vbCrLf
    BufStr = BufStr & "GetTempFileName()" & vbTab
    BufStr = BufStr & "GetSpecialFolder()" & vbTab
    BufStr = BufStr & "FormatFileSize()" & vbCrLf
    BufStr = BufStr & "GetFile()" & vbTab & vbTab
    BufStr = BufStr & "FileInUse()" & vbTab & vbTab
    BufStr = BufStr & "LocalDateToUTC()" & vbCrLf
    BufStr = BufStr & "UTCToLocalDate()" & vbCrLf
    BufStr = BufStr & "**** 46 functions found" & vbCrLf
    BufStr = BufStr & "-------------------------------------------------------------------------" & vbCrLf
    BufStr = BufStr & "Contact the author" & vbCrLf
    BufStr = BufStr & "marclei@spnorte.com" & vbCrLf
    Text1.Text = BufStr
End Sub

Private Sub DebugOut(ByVal Text As String, Optional TabCount As Integer = 4)
    Dim Size As Long
    
    Text = Space(TabCount) & Text & vbCrLf
    With Text1
        Size = Len(.Text)
        .SelStart = Size
        .SelLength = Len(Text)
        .SelText = Text
    End With
End Sub

Private Sub Form_Load()
    ShowReadMe
End Sub

Private Sub ShowAbout()
    Dim BufStr As String
    
    BufStr = BufStr & "FileSystem Module" & vbCrLf
    BufStr = BufStr & "(february, 04 2001)" & vbCrLf
    BufStr = BufStr & "-------------------------------------------------------------------------" & vbCrLf
    BufStr = BufStr & "Created by  Marclei V Silva" & vbCrLf
    BufStr = BufStr & "Description Module that re-writes several FileSystemObject functions" & vbCrLf
    BufStr = BufStr & "" & vbCrLf
    BufStr = BufStr & "  Alot of the information contained inside this file was originally" & vbCrLf
    BufStr = BufStr & "  obtained from several authors on the net and most of it has since been" & vbCrLf
    BufStr = BufStr & "  modified in some way." & vbCrLf
    BufStr = BufStr & "" & vbCrLf
    BufStr = BufStr & "Disclaimer: This file is public domain, updated periodically by" & vbCrLf
    BufStr = BufStr & "  Marclei, (marclei@spnorte.com), Use it at your own risk." & vbCrLf
    BufStr = BufStr & "  Neither myself(marclei) or anyone related to spnorte.com" & vbCrLf
    BufStr = BufStr & "  may be held liable for its use, or misuse." & vbCrLf
    BufStr = BufStr & "" & vbCrLf
    BufStr = BufStr & "Declare check Jan 29, 2001. (Marclei, marclei@spnorte.com)" & vbCrLf
    BufStr = BufStr & "  Works fine running on windows NT 4.0, but I have to check" & vbCrLf
    BufStr = BufStr & "  Win 9x platform. This release I am not handling NT security" & vbCrLf
    BufStr = BufStr & "  concerning register values or file access, this is something" & vbCrLf
    BufStr = BufStr & "  I am working on." & vbCrLf
    BufStr = BufStr & "" & vbCrLf
    BufStr = BufStr & "Declare check Feb 04, 2001. (Marclei, marclei@spnorte.com)" & vbCrLf
    BufStr = BufStr & "  First release with 46 public functions and routines" & vbCrLf
    BufStr = BufStr & "" & vbCrLf
    BufStr = BufStr & "NOTES:" & vbCrLf
    BufStr = BufStr & "  (1) Many of these functions and procedures have not been tested hard" & vbCrLf
    BufStr = BufStr & "      so if you find any bug, please send them to marclei@spnorte.com and this" & vbCrLf
    BufStr = BufStr & "      module will be updated and reposted. Thanks!" & vbCrLf
    BufStr = BufStr & "  (2) These functions are not so robust as FileSystemObject but to" & vbCrLf
    BufStr = BufStr & "      acomplish small tasks it is very useful" & vbCrLf
    BufStr = BufStr & "-------------------------------------------------------------------------" & vbCrLf & vbCrLf
    BufStr = BufStr & "Contact the author" & vbCrLf
    BufStr = BufStr & "marclei@spnorte.com" & vbCrLf
    Text1.Text = BufStr
End Sub

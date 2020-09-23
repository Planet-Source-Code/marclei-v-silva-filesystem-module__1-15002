Attribute VB_Name = "FileSystem"
' ******************************************************************************
' Module        : FileSystem
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 12:52:59
' Description   : Module that re-writes several FileSystemObject functions
'
'   Alot of the information contained inside this file was originally
'   obtained from several authors on the net and most of it has since been
'   modified in some way.
'
' Disclaimer: This file is public domain, updated periodically by
'   Marclei, (marclei@spnorte.com), Use it at your own risk.
'   Neither myself(marclei) or anyone related to spnorte.com
'   may be held liable for its use, or misuse.
'
' Declare check Jan 29, 2001. (Marclei, marclei@spnorte.com)
'   Works fine running on windows NT 4.0, but I have to check
'   Win 9x platform. This release I am not handling NT security
'   concerning register values or file access, this is something
'   I am working on.
'
' Declare check Feb 04, 2001. (Marclei, marclei@spnorte.com)
'   First release with 46 public functions and routines
'
' NOTES:
'   (1) Many of these functions and procedures have not been tested hard
'       so if you find any bug, please send them to marclei@spnorte.com and this
'       module will be updated and reposted. Thanks!
'   (2) These functions are not so robust as FileSystemObject but to
'       acomplish small tasks it is very useful
' ******************************************************************************
Option Explicit
Option Compare Text
DefLng A-Z

' Keep up with the errors
Const g_ErrConstant As Long = vbObjectError + 1000
Const m_constClassName = "FileSystem"

Private m_lngErrNum As Long
Private m_strErrStr As String
Private m_strErrSource As String

' registry constants
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

' registry specific access rights
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = &H3F

' open/create Options
Public Const REG_OPTION_NON_VOLATILE = 0&
Public Const REG_OPTION_VOLATILE = &H1

' key creation/open disposition
Public Const REG_CREATED_NEW_KEY = &H1
Public Const REG_OPENED_EXISTING_KEY = &H2

' masks for the predefined standard access types
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const SPECIFIC_RIGHTS_ALL = &HFFFF

' define severity codes
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_ACCESS_DENIED = 5
Private Const ERROR_NO_MORE_ITEMS = 259

' misc error
Private Const ERROR_SHARING_VIOLATION As Long = 32

' structure to handle picture information
Public Type TPictureInfo
    Width As Long
    Height As Long
    ColorDepth As Double
    Type As String
End Type

Public Const OFS_MAXPATHNAME = 260
Public Const OF_READWRITE = &H2
Public Const OF_CREATE = &H1000
Public Const OF_READ = &H0
Public Const OF_WRITE = &H1

Public Const INVALID_HANDLE_VALUE = -1
Public Const MAX_PATH = 260

' flag Attributes
Public Enum EFileAttributes
     FILE_ATTRIBUTE_READONLY = &H1
     FILE_ATTRIBUTE_HIDDEN = &H2
     FILE_ATTRIBUTE_SYSTEM = &H4
     FILE_ATTRIBUTE_DIRECTORY = &H10
     FILE_ATTRIBUTE_ARCHIVE = &H20
     FILE_ATTRIBUTE_NORMAL = &H80
     FILE_ATTRIBUTE_TEMPORARY = &H100
     FILE_ATTRIBUTE_COMPRESSED = &H800
End Enum

' enumerate special folders
Public Enum ESpecialFolders
    WindowsFolder = 0
    SystemFolder = 1
    TemporaryFolder = 2
End Enum

' Creation Disposition
Public Enum EFileCreationDisposition
    CREATE_ALWAYS = 2
    CREATE_NEW = 1
    OPEN_ALWAYS = 4
    OPEN_EXISTING = 3
    TRUNCATE_EXISTING = 5
End Enum

' file flags
Public Enum EFileFlags
     FILE_FLAG_BACKUP_SEMANTICS = &H2000000
     FILE_FLAG_DELETE_ON_CLOSE = &H4000000
     FILE_FLAG_NO_BUFFERING = &H20000000
     FILE_FLAG_OVERLAPPED = &H40000000
     FILE_FLAG_POSIX_SEMANTICS = &H1000000
     FILE_FLAG_RANDOM_ACCESS = &H10000000
     FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
     FILE_FLAG_WRITE_THROUGH = &H80000000
End Enum

' move file constants
Private Const MOVEFILE_COPY_ALLOWED = &H2
Private Const MOVEFILE_DELAY_UNTIL_REBOOT = &H4
Private Const MOVEFILE_REPLACE_EXISTING = &H1
Private Const MOVEFILE_WRITE_THROUGH = &H8

' file info constants
Private Const SHGFI_ICON = &H100                         '  get icon
Private Const SHGFI_DISPLAYNAME = &H200                  '  get display name
Private Const SHGFI_TYPENAME = &H400                     '  get type name
Private Const SHGFI_ATTRIBUTES = &H800                   '  get attributes
Private Const SHGFI_ICONLOCATION = &H1000                '  get icon location
Private Const SHGFI_EXETYPE = &H2000                     '  return exe type
Private Const SHGFI_SYSICONINDEX = &H4000                '  get system icon index
Private Const SHGFI_LINKOVERLAY = &H8000                 '  put a link overlay on icon
Private Const SHGFI_SELECTED = &H10000                   '  show icon in selected state
Private Const SHGFI_LARGEICON = &H0                      '  get large icon
Private Const SHGFI_SMALLICON = &H1                      '  get small icon
Private Const SHGFI_OPENICON = &H2                       '  get open icon
Private Const SHGFI_SHELLICONSIZE = &H4                  '  get shell size icon
Private Const SHGFI_PIDL = &H8                           '  pszPath is a pidl
Private Const SHGFI_USEFILEATTRIBUTES = &H10             '  use passed dwFileAttribute

' file IO modes
Public Enum EIOModes
    ForReading = 1
    ForWriting = 2
    ForAppending = 3
End Enum

Public Const GENERIC_ALL = &H10000000
Public Const GENERIC_EXECUTE = &H20000000
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000

' file sharing constants
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2

Private Const MAXDWORD As Long = &HFFFF

Private Type SHFILEINFO
    hIcon As Long                      '  out: icon
    iIcon As Long                      '  out: icon index
    dwAttributes As Long               '  out: SFGAO_ flags
    szDisplayName As String * MAX_PATH '  out: display name (or path)
    szTypeName As String * 80          '  out: type name
End Type

Private Type OFSTRUCT
    cBytes      As Byte
    fFixedDisk  As Byte
    nErrCode    As Integer
    Reserved1   As Integer
    Reserved2   As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Private Type FILETIME
    dwLowDateTime     As Long
    dwHighDateTime    As Long
End Type

Private Type SYSTEMTIME
    wYear          As Integer
    wMonth         As Integer
    wDayOfWeek     As Integer
    wDay           As Integer
    wHour          As Integer
    wMinute        As Integer
    wSecond        As Integer
    wMilliseconds  As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

' custom udt for holding file information
' from FindFirst() and FindNext() function
Public Type TFile
    Attributes As Long
    DateCreated As Date
    DateLastAccessed As Date
    DateLastModified As Date
    Size As Long
    Alternate As String
    Name As String
    Path As String
    TypeName As String
    Directory As Boolean
    ParentFolder As String
    DisplayName As String
    Extension As String
    ShortName As String
    ShortPathName As String
    FullPathName As String
    CompressedSize As Long
    hIcon As Long
End Type

Private Type TIME_ZONE_INFORMATION
   Bias As Long
   StandardName(32) As Integer
   StandardDate As SYSTEMTIME
   StandardBias As Long
   DaylightName(32) As Integer
   DaylightDate As SYSTEMTIME
   DaylightBias As Long
End Type

' security descriptor
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Enum ERegValueTypes
    REG_NONE = 0&
    REG_SZ = 1&
    REG_EXPAND_SZ = 2&
    REG_BINARY = 3&
    REG_DWORD = 4&
    REG_DWORD_LITTLE_ENDIAN = 4&
    REG_DWORD_BIG_ENDIAN = 5&
    REG_LINK = 6&
    REG_MULTI_SZ = 7&
    REG_RESOURCE_LIST = 8&
    REG_FULL_RESOURCE_DESCRIPTOR = 9&
    REG_RESOURCE_REQUIREMENTS_LIST = 10&
End Enum

' used to get error messages directly from the
' system instead of hard-coding them
Private Const FORMAT_MESSAGE_FROM_SYSTEM     As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS  As Long = &H200

' find file functions
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

' misc functions
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetFileType Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function GetCompressedFileSize Lib "kernel32" Alias "GetCompressedFileSizeA" (ByVal lpFileName As String, lpFileSizeHigh As Long) As Long

' date and time function
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long

' special folders functions
Private Declare Function ApiWindDir Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function ApiSysDir Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

' temporary file functions
Private Declare Function ApiGetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function ApiGetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

' file and folder operations
Private Declare Function ApiSetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function ApiGetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, lpFilePart As Long) As Long
Private Declare Function ApiGetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal nBufferLength As Long) As Long
Private Declare Function ApiGetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function ApiSetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function ApiOpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function ApiCloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hFile As Long) As Long
Private Declare Function ApiCopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function ApiDeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function ApiMoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Declare Function ApiCreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ApiMoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long
Private Declare Function ApiCreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare Function ApiRemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long

' shell32 functions
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

' registry Function Prototypes
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal szData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

' error handling functions
' Most of the routines here return a boolean
' indicating success or not. But you can provide
' error handling by using the api functions below
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

' Utilities function

' ******************************************************************************
' Routine       : GetRegValue
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 12:52:59
' Description   : Get a registry value
' Inputs        :
' Outputs       :
' Credits       : LocalWeb Server Project (do not know the author name)
' Modifications : Added the key type to be retrieved
' Remarks       :
' ******************************************************************************
Private Function GetRegValue(hKey As Long, lpszSubKey As String, szKey As String, hValueType As ERegValueTypes, szDefault As String) As Variant
    On Error GoTo ErrorRoutineErr:
    
    Dim phkResult As Long
    Dim lResult As Long
    Dim szBuffer As String
    Dim lBuffSize As Long
    
    ' create Buffer
    szBuffer = Space(255)
    lBuffSize = Len(szBuffer)
    ' open the key
    RegOpenKeyEx hKey, lpszSubKey, 0, 1, phkResult
    ' query the value
    lResult = RegQueryValueEx(phkResult, szKey, 0, hValueType, szBuffer, lBuffSize)
    ' close the key
    RegCloseKey phkResult
    ' return obtained value
    If lResult = ERROR_SUCCESS Then
        GetRegValue = Left(szBuffer, lBuffSize - 1)
    Else
        GetRegValue = szDefault
    End If
    Exit Function
    
ErrorRoutineErr::
    GetRegValue = ""
End Function


' ******************************************************************************
' Routine       : ConcatString
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 13:40:31
' Description   : Concatenate two strings with a separator string
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Private Function ConcatString(ByVal s1 As String, ByVal s2 As String, sep As String) As String
    If s1 = "" Then
        s1 = s2
    ElseIf Trim(s2) <> "" Then
        s1 = s1 & sep & s2
    End If
    ConcatString = s1
End Function

' ******************************************************************************
' Routine       : GetFileDateString
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 04/02/01 10:52:00
' Description   : Returns a string based on the file time
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Private Function GetFileDateString(CT As FILETIME) As String
    Dim ST As SYSTEMTIME
    Dim ds As Single
    
    ' convert the passed FILETIME to a valid SYSTEMTIME format for display
    If FileTimeToSystemTime(CT, ST) Then
        On Error Resume Next
        GetFileDateString = DateSerial(ST.wYear, ST.wMonth, ST.wDay) & " " & TimeSerial(ST.wHour, ST.wMinute, ST.wSecond)
        Else: GetFileDateString = ""
    End If
End Function

' ******************************************************************************
' Routine       : GetSystemDateString
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 04/02/01 10:52:31
' Description   : returns a string based on an system time
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Private Function GetSystemDateString(ST As SYSTEMTIME) As String
    Dim ds As Single
  
    ds = DateSerial(ST.wYear, ST.wMonth, ST.wDay)
    If ds Then
        GetSystemDateString = DateSerial(ST.wYear, ST.wMonth, ST.wDay) & " " & TimeSerial(ST.wHour, ST.wMinute, ST.wSecond)
    Else
        GetSystemDateString = ""
    End If
End Function

' ******************************************************************************
' Routine       : TrimNull
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 04/02/01 10:52:47
' Description   : Remove padding null characters from a string
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Private Function TrimNull(OriginalStr As String) As String
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    TrimNull = OriginalStr
End Function

'--------------------------------------
' file system core functions
'--------------------------------------

' ******************************************************************************
' Routine       : StripBkSlash
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 04/02/01 10:53:16
' Description   : Remove the back slash from the path given
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Public Function StripBkSlash(ByVal Path As String) As String
    If Right(Path, 1) = "\" Or Right(Path, 1) = "/" Then
        Path = Mid(Path, 1, Len(Path) - 1)
    End If
    StripBkSlash = Path
End Function

' ******************************************************************************
' Routine       : StripLtSlash
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 04/02/01 10:54:05
' Description   : Remove the left slash from the path given
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Public Function StripLtSlash(ByVal Path As String) As String
    If Left(Path, 1) = "\" Or Left(Path, 1) = "/" Then
        Path = Mid(Path, 2)
    End If
    StripLtSlash = Path
End Function

' ******************************************************************************
' Routine       : AddBkSlash
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 04/02/01 10:54:35
' Description   : Add path back slash
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Public Function AddBkSlash(ByVal Path As String) As String
    If Right(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    AddBkSlash = Path
End Function

' ******************************************************************************
' Routine       : AddLtSlash
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 01/02/01 2:41:32
' Description   : Add a left slash to the file path
' Inputs        : Path :  path we wish to add a left slash
' Outputs       :
' Credits       :
' Modifications :
' Remarks       : Ex.: AddLtSlash("Windows\System")
'                      returns \Windows\System
' ******************************************************************************
Public Function AddLtSlash(ByVal Path As String) As String
    If Left(Path, 1) <> "\" Then
        Path = "\" & Path
    End If
    AddLtSlash = Path
End Function

' ******************************************************************************
' Routine       : SetCurrentDirectory
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 01/02/01 2:43:48
' Description   : Set windows current directory
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Public Sub SetCurrentDirectory(Path As String)
    ApiSetCurrentDirectory Path
End Sub

' ******************************************************************************
' Routine       : GetParentFolderName
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 10:51:42
' Description   : Strip off the last path name to get the parent folder's name
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' ******************************************************************************
Public Function GetParentFolderName(ByVal Path As String) As String
    Dim Pos As Integer

    Pos = InStrRev(Path, "\")
    If Pos > 0 Then
        Path = Left(Path, Pos - 1)
    End If
    GetParentFolderName = Path
End Function

' ******************************************************************************
' Routine       : GetFileAttributes
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 9:27:49
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : Get file attributes
'                 VB has GetAttr() function and this one makes
'                 the same thing, it was put here just for
'                 compatibility
' ******************************************************************************
Public Function GetFileAttributes(ByVal Filename As String) As EFileAttributes
    GetFileAttributes = ApiGetFileAttributes(Filename)
End Function

' ******************************************************************************
' Routine       : SetFileAttributes
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 9:29:05
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : Set file attributes
'                 VB has SetAttr() function and this one makes
'                 the same thing, it was put here just for
'                 compatibility
' ******************************************************************************
Public Function SetFileAttributes(ByVal Filename As String, ByVal FileAttributes As EFileAttributes) As EFileAttributes
    SetFileAttributes = ApiSetFileAttributes(Filename, FileAttributes)
End Function

' ******************************************************************************
' Routine       : FileExists
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/019:30:15
' Description   : Check the existence of a file
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' ******************************************************************************
Public Function FileExists(Filename As String) As Boolean
    Dim File As WIN32_FIND_DATA
    Dim hFile As Long
   
    hFile = FindFirstFile(Filename, File)
    FileExists = (hFile <> INVALID_HANDLE_VALUE)
    FindClose hFile
End Function

' ******************************************************************************
' Routine       : FolderExists
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 9:30:32
' Description   : Check the existence of folder
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' ******************************************************************************
Public Function FolderExists(PathSpec As String) As Boolean
    Dim hFile As Long
    Dim File As WIN32_FIND_DATA
    Dim FilePath As String
    
    ' remove training slash before verifying
    FilePath = StripBkSlash(PathSpec)
    ' call the API pasing the folder
    hFile = FindFirstFile(FilePath, File)
    ' if a valid file handle was returned,
    ' and the directory attribute is set
    ' the folder exists
    FolderExists = (hFile <> INVALID_HANDLE_VALUE) And _
       (File.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY)
    ' clean up
    FindClose hFile
End Function

' ******************************************************************************
' Routine       : GetFileName
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 9:30:54
' Description   : Returns the last component of specified path that is not part
'                 of the drive specification.
' Inputs        : FilePath          Required. The path (absolute or relative) to
'                                   a specific file.
' Outputs       :
' Credits       : Extracted from Microsoft setup wizard "common.bas" module
' Modifications : Changed the name from "ExtractFile" to "GetFileName"
' ******************************************************************************
Public Function GetFileName(FilePath As String) As String
    Dim r As Long
    Dim strPathPart As String
    Dim strTempPath As String
    Dim lngOriginalLength As Long
    Dim lngFinalLength As Long
    
    strTempPath = FilePath
    lngOriginalLength = Len(strTempPath)
    Do While InStr(1, strTempPath, "\", vbTextCompare) <> 0
        r = InStr(1, strTempPath, "\", vbTextCompare)
        strTempPath = Right(strTempPath, ((Len(strTempPath) - r)))
    Loop
    GetFileName = strTempPath
End Function

' ******************************************************************************
' Routine       : GetBasename
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 13:59:09
' Description   : Returns a string containing the base name of the last component
'                 less any file extension, in a path.
' Inputs        : Path : Required. The path specification for the component whose
'                        base name is to be returned.
' Outputs       :
' Credits       :
' Modifications :
' Remarks       : The GetBaseName method returns a zero-length string ("") if no
'                 component matches the path argument.
' ******************************************************************************
Public Function GetBasename(Path As String) As String
    Dim Filename As String
    Dim DotPos As Integer
    
    Filename = GetFileName(Path)
    DotPos = InStrRev(Filename, ".")
    If DotPos > 0 Then
        Filename = Mid(Filename, 1, DotPos - 1)
    End If
    GetBasename = Filename
End Function

' ******************************************************************************
' Routine       : GetPathName
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 9:33:33
' Inputs        : FilePath : full file path
' Outputs       : File path
' Credits       : Extracted from Microsoft setup wizard "common.bas" module
' Modifications :
' Description   : Extracts a path from its path & file
' ******************************************************************************
Public Function GetPathName(FilePath As String) As String
    Dim intPos As Long
    Dim i As Long

    For i = Len(FilePath) To 1 Step -1
        If Mid(FilePath, i, 1) = "\" Then
            GetPathName = Left(FilePath, i - 1)
            Exit Function
        End If
    Next i
    GetPathName = FilePath
End Function

' ******************************************************************************
' Routine       : GetFileInfo
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 9:35:37
' Description   : Get extended file information
' Inputs        : Path  Parent folder
'                 WFD   Windows find information
'                 File  File extended info
' Outputs       :
' Credits       :
' Modifications :
' ******************************************************************************
Private Sub GetFileInfo(Path As String, WFD As WIN32_FIND_DATA, File As TFile)
    Dim FI As SHFILEINFO
    Dim Buffer As String
    
    File.Directory = (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY)
    File.ParentFolder = StripBkSlash(Path)
    On Error Resume Next
    File.Name = Trim(TrimNull(WFD.cFileName))
    File.Extension = GetExtensionName(File.Name)
    ' short name same as long, if cAlternate element empty.
    If InStr(WFD.cAlternate, vbNullChar) = 1 Then
        File.ShortName = UCase(File.Name)
    Else
        File.ShortName = TrimNull(WFD.cAlternate)
    End If
    File.Path = Path & "\" & File.Name
    File.ShortPathName = GetShortPathName(File.Path)
    File.FullPathName = GetFullPathName(File.Path)
    ' Retrieve compressed size.
    File.CompressedSize = GetCompressedFileSize(File.Path, 0&)
    File.DateLastModified = GetFileDateString(WFD.ftLastWriteTime)
    File.DateCreated = GetFileDateString(WFD.ftCreationTime)
    File.DateLastAccessed = GetFileDateString(WFD.ftLastAccessTime)
    File.Attributes = WFD.dwFileAttributes
    File.Size = (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
    ' get extended file information
    SHGetFileInfo File.Path, 0, FI, Len(FI), SHGFI_ICON Or SHGFI_DISPLAYNAME Or SHGFI_TYPENAME
    File.DisplayName = TrimNull(FI.szDisplayName)
    File.TypeName = TrimNull(FI.szTypeName)
    File.hIcon = FI.hIcon
    ' Confirm displayable typename.
    If Trim(File.TypeName) = "" Then
        File.TypeName = Trim(UCase(File.Extension) & " File")
    End If
End Sub

' ******************************************************************************
' Routine       : CreateFolder
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 27/01/01 21:53:29
' Description   : Creates nested directories on the drive
'                 included in the path by parsing the final
'                 directory string into a directory array,
'                 and looping through each to create the final path.
'                 The path could be passed to this method as a
'                 pre-filled array, reducing the code.
' Inputs        :
' Outputs       :
' Credits       : Was extract from a Vbnet web site sample app
' Modifications : Some variables were renamed to maintaim module
'                 standards
' ******************************************************************************
Public Function CreateFolder(PathSpec As String) As Boolean
    Dim r As Long
    Dim Sa As SECURITY_ATTRIBUTES
    Dim drivePart As String
    Dim newDirectory  As String
    Dim Item As String
    Dim sfolders() As String
    Dim Pos As Integer
    Dim x As Integer
    Dim retVal As Long
    Dim CompleteDirectory As String
    
    CompleteDirectory = PathSpec
    ' must have a trailing slash for
    ' the GetPart routine below
    If Right$(CompleteDirectory, 1) <> "\" Then
        CompleteDirectory = CompleteDirectory & "\"
    End If
    ' if there is a drive in the string, get it
    ' else, just use nothing - assumes current drive
    Pos = InStr(CompleteDirectory, ":")
    If Pos Then
        drivePart = GetPart(CompleteDirectory, "\")
        Else: drivePart = ""
    End If
    ' now get the rest of the items that
    ' make up the string
    Do Until CompleteDirectory = ""
        ' strip off one item (i.e. "Files\")
        Item = GetPart(CompleteDirectory, "\")
        ' add it to an array for later use, and
        ' if this is the first item (x=0),
        ' append the drivepart
        ReDim Preserve sfolders(0 To x) As String
        If x = 0 Then Item = drivePart & Item
        sfolders(x) = Item
        ' increment the array counter
        x = x + 1
    Loop
    ' Now create the directories.
    ' Because the first directory is
    ' 0 in the array, reinitialize x to -1
    x = -1
    Do
        x = x + 1
        ' just keep appending the folders in the
        ' array to newDirectory.  When x=0 ,
        ' newDirectory is "", so the
        ' newDirectory gets assigned drive:\firstfolder.
        ' Subsequent loops adds the next member of the
        ' array to the path, forming a fully qualified
        ' path to the new directory.
        newDirectory = newDirectory & sfolders(x)
        If FolderExists(newDirectory) = False Then
            ' the only member of the SA type needed (on
            ' a win95/98 system at least)
            Sa.nLength = LenB(Sa)
            retVal = ApiCreateDirectory(newDirectory, Sa)
            If retVal = ERROR_SUCCESS Then
                CreateFolder = False
                Exit Function
            End If
        End If
    Loop Until x = UBound(sfolders)
    CreateFolder = True
End Function

' ******************************************************************************
' Routine       : GetPart
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 27/01/01 21:52:58
' Description   : Takes a string separated by "delimiter",
'                 splits off 1 item, and shortens the string
'                 so that the next item is ready for removal.
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' ******************************************************************************
Private Function GetPart(startStrg As String, Delimiter As String) As String
    Dim c As Integer
    Dim Item As String
    c = 1
    Do
        If Mid$(startStrg, c, 1) = Delimiter Then
            Item = Mid$(startStrg, 1, c)
            startStrg = Mid$(startStrg, c + 1, Len(startStrg))
            GetPart = Item
            Exit Function
        End If
        c = c + 1
    Loop
End Function

' ******************************************************************************
' Routine       : HasWildcards
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 04/02/01 11:31:05
' Description   : Detects if string has wild cards info
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Private Function HasWildcards(S As String) As Boolean
    HasWildcards = InStr(S, "?") Or InStr(S, "*")
End Function

' ******************************************************************************
' Routine       : DeleteFile
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 9:50:09
' Description   : Delete a specified file
' Inputs        :
'                 filespec Required.    The name of the file to delete. The filespec
'                                       can contain wildcard characters in the last
'                                       path component.
'                 force Optional.       Boolean value that is True if files with the
'                                       read-only attribute set are to be deleted;
'                                       False (default) if they are not.
' Outputs       :
' Credits       :
' Modifications :
' ******************************************************************************
Public Function DeleteFile(ByVal FileSpec As String, Optional Force As Boolean = False) As Boolean
    ' if filespec has wildcards then call the
    ' DeleteFileList() to delete mass files
    ' else simply call DeleteSingleFile() function
    If HasWildcards(FileSpec) Then
        DeleteFile = DeleteFileList(FileSpec, Force)
    Else
        DeleteFile = DeleteSingleFile(FileSpec, Force)
    End If
End Function

Private Function DeleteSingleFile(ByVal FileSpec As String, Optional Force As Boolean = False) As Boolean
    Dim lngRetVal As Long
    Dim Attr As EFileAttributes
    
    If FileExists(FileSpec) Then
        Attr = GetFileAttributes(FileSpec)
        ' if force was set to true we have to change
        ' file read-only atribute so the file can be
        ' deleted else the function will fail
        If ((Attr And FILE_ATTRIBUTE_READONLY) And (Force = True)) Then
            SetFileAttributes FileSpec, FILE_ATTRIBUTE_NORMAL
        End If
        lngRetVal = ApiDeleteFile(FileSpec)
        DeleteSingleFile = (lngRetVal <> ERROR_SUCCESS)
    End If
End Function

Private Function DeleteFileList(ByVal FileSpec As String, Optional Force As Boolean = False) As Boolean
    Dim File As WIN32_FIND_DATA
    Dim SourceFile As String
    Dim Filename As String
    Dim hSearch As Long
    Dim FilePath As String
    Dim Source As String
    
    hSearch = FindFirstFile(FileSpec, File)
    Source = GetPathName(FileSpec)
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do
            ' get the file name
            Filename = TrimNull(File.cFileName)
            ' override current and parent path
            If (Filename <> ".") And (Filename <> "..") Then
                ' this must be a file
                If Not (File.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                    ' get the source file
                    SourceFile = AddBkSlash(Source) & Filename
                    ' attempt to copy the file
                    If DeleteSingleFile(SourceFile, Force) = False Then
                        ' delete failed then exit
                        GoTo Err_DeleteFileList
                    End If
                End If
            End If
        Loop While FindNextFile(hSearch, File)
        ' must do this
        FindClose hSearch
        ' delete was succesful
        DeleteFileList = True
    End If
    
    Exit Function
Err_DeleteFileList:
    If hSearch <> INVALID_HANDLE_VALUE Then
        FindClose hSearch
    End If
    DeleteFileList = False
End Function

' ******************************************************************************
' Routine       : CopyFile
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 10:05:34
' Description   : Copies one or more files from one location to another.
' Inputs        :
'                 source Required.      Character string file specification, which
'                                       can include wildcard characters, for one
'                                       or more files to be copied.
'                 destination Required. Character string destination where the file
'                                       or files from source are to be copied.
'                                       Wildcard characters are not allowed.
'                 overwrite Optional.   Boolean value that indicates if existing
'                                       files are to be overwritten. If True, files
'                                       are overwritten; if False, they are not.
'                                       The default is True. Note that CopyFile will
'                                       fail if destination has the read-only
'                                       attribute set, regardless of the value of overwrite.
' Outputs       : Boolean indicating copy success
' Credits       :
' Modifications :
' ******************************************************************************
Public Function CopyFile(ByVal Source As String, ByVal Destination As String, Optional ByVal OverWriteFiles As Boolean) As Boolean
    ' if source has wildcards then call the
    ' CopyFileList() to delete mass files
    ' else simply call CopySingleFile() function
    If HasWildcards(Source) Then
        CopyFile = CopyFileList(Source, Destination, OverWriteFiles)
    Else
        CopyFile = CopySingleFile(Source, Destination, OverWriteFiles)
    End If
End Function

Private Function CopySingleFile(ByVal Source As String, ByVal Destination As String, Optional ByVal OverWriteFiles As Boolean) As Boolean
    Dim lngRetVal As Long
    Dim Filename As String
    
    ' check we are copying to a file or folder
    ' to copy to a folder, destination must ends
    ' with a path separator or else this function will fail
    If Right(Trim(Destination), 1) = "\" Then
        ' extract source file name
        Filename = GetFileName(Source)
        ' apply it to destination
        Destination = Destination & Filename
    End If
    lngRetVal = ApiCopyFile(Source, Destination, Not OverWriteFiles)
    CopySingleFile = (lngRetVal <> ERROR_SUCCESS)
End Function
    
Private Function CopyFileList(ByVal Source As String, ByVal Destination As String, Optional ByVal OverWriteFiles As Boolean) As Boolean
    Dim File As WIN32_FIND_DATA
    Dim SourceFile As String
    Dim Filename As String
    Dim hSearch As Long
    
    hSearch = FindFirstFile(Source, File)
    Source = GetPathName(Source)
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do
            ' get the file name
            Filename = TrimNull(File.cFileName)
            ' override current and parent path
            If (Filename <> ".") And (Filename <> "..") Then
                ' this must be a file not a directory
                If Not (File.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                    ' get the source file
                    SourceFile = AddBkSlash(Source) & Filename
                    ' attempt to copy the file
                    If CopySingleFile(SourceFile, Destination, OverWriteFiles) = False Then
                        ' copy failed to complete
                        GoTo Err_CopyFileList
                    End If
                End If
            End If
        Loop While FindNextFile(hSearch, File)
        ' must to do this
        FindClose hSearch
        ' copy was succesful
        CopyFileList = True
    End If
    
    Exit Function
Err_CopyFileList:
    If hSearch <> INVALID_HANDLE_VALUE Then
        FindClose hSearch
    End If
    CopyFileList = False
End Function

' ******************************************************************************
' Routine       : MoveFile
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 10:53:43
' Description   : Moves one or more files from one location to another
' Inputs        :
'                 source Required.      The path to the file or files to be moved.
'                                       The source argument string can contain
'                                       wildcard characters in the last path
'                                       component only.
'                 destination Required. The path where the file or files are to
'                                       be moved. The destination argument can't
'                                       contain wildcard characters.
' Outputs       :
' Credits       :
' Modifications :
' ******************************************************************************
Public Function MoveFile(ByVal Source As String, ByVal Destination As String) As Boolean
    ' if source has wildcards then call the
    ' MoveFileList() to delete mass files
    ' else simply call MoveSingleFile() function
    If HasWildcards(Source) Then
        MoveFile = MoveFileList(Source, Destination)
    Else
        MoveFile = MoveSingleFile(Source, Destination)
    End If
End Function

Private Function MoveSingleFile(ByVal Source As String, ByVal Destination As String) As Boolean
    Dim lngRetVal As Long
    Dim Filename As String
    
    ' check we are copying to a file or folder
    ' to copy to a folder, destination must have a trailing back
    ' slash or else this function will fail
    If Right(Trim(Destination), 1) = "\" Then
        ' extract source file name
        Filename = GetFileName(Source)
        ' apply it to destination
        Destination = Destination & Filename
    End If
    lngRetVal = ApiMoveFileEx(Source, Destination, MOVEFILE_REPLACE_EXISTING Or MOVEFILE_COPY_ALLOWED Or MOVEFILE_WRITE_THROUGH)
    MoveSingleFile = (lngRetVal <> ERROR_SUCCESS)
End Function

Private Function MoveFileList(ByVal Source As String, ByVal Destination As String) As Boolean
    Dim File As WIN32_FIND_DATA
    Dim SourceFile As String
    Dim DestFile As String
    Dim Filename As String
    Dim hSearch As Long
    Dim bCreateFile As Boolean
    
    hSearch = FindFirstFile(Source, File)
    Source = GetPathName(Source)
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do
            ' get file name from file data
            Filename = TrimNull(File.cFileName)
            ' skip parent and current path
            If (Filename <> ".") And (Filename <> "..") Then
                ' this must be a file not a directory
                If Not (File.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                    ' build source file
                    SourceFile = AddBkSlash(Source) & Filename
                    ' try to move the file(s)
                    If MoveSingleFile(SourceFile, Destination) = False Then
                        ' move failed then exit
                        GoTo Err_MoveFileList
                    End If
                End If
            End If
        Loop While FindNextFile(hSearch, File)
        ' must close find handle
        FindClose hSearch
        ' move succeed
        MoveFileList = True
    End If
    
    Exit Function
Err_MoveFileList:
    If hSearch <> INVALID_HANDLE_VALUE Then
        FindClose hSearch
    End If
    MoveFileList = False
End Function

' ******************************************************************************
' Routine       : CopyFolder
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 10:05:34
' Description   : Copies one or more folders from one location to another.
' Inputs        :
'                 source Required.      Character string folder specification, which
'                                       can include wildcard characters, for one
'                                       or more folder to be copied.
'                 destination Required. Character string destination where the folder
'                                       or folders from source are to be copied.
'                                       Wildcard characters are not allowed.
'                 overwrite Optional.   Boolean value that indicates if existing
'                                       folders are to be overwritten. If True, folders
'                                       are overwritten; if False, they are not.
'                                       The default is True. Note that CopyFolder will
'                                       fail if destination has the read-only
'                                       attribute set, regardless of the value of overwrite.
' Outputs       : Boolean indicating copy success
' Credits       :
' Modifications :
' ******************************************************************************
Public Function CopyFolder(ByVal Source As String, ByVal Destination As String, Optional ByVal OverWriteFiles As Boolean) As Boolean
    Dim DestPath As String
    
    ' if source has wildcards then call the
    ' CopyFolderList() to delete mass files
    ' else simply call CopySingleFolder() function
    If HasWildcards(Source) Then
        CopyFolder = CopyFolderList(Source, Destination, OverWriteFiles)
    Else
        CopyFolder = CopySingleFolder(Source, Destination, OverWriteFiles)
    End If
End Function

Private Function CopySingleFolder(ByVal Source As String, ByVal Destination As String, Optional ByVal OverWriteFiles As Boolean) As Boolean
    Dim DestPath As String
    
    If Right(Destination, 1) <> "\" Then
        If CreateFolder(Destination) = False Then
            Exit Function
        End If
        Destination = AddBkSlash(Destination)
    End If
    DestPath = GetFileName(Source)
    DestPath = Destination & DestPath
    If CreateFolder(DestPath) Then
        CopySingleFolder = CopyFolderList(AddBkSlash(Source) & "*.*", DestPath, OverWriteFiles)
    End If
End Function
    
Private Function CopyFolderList(ByVal Source As String, ByVal Destination As String, Optional ByVal OverWriteFiles As Boolean) As Boolean
    Dim File As WIN32_FIND_DATA
    Dim SourceFile As String
    Dim DestFile As String
    Dim Filename As String
    Dim hSearch As Long
    
    hSearch = FindFirstFile(Source, File)
    Source = GetPathName(Source)
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do
            ' get the file name
            Filename = TrimNull(File.cFileName)
            ' override current and parent path
            If (Filename <> ".") And (Filename <> "..") Then
                ' get the source file
                SourceFile = AddBkSlash(Source) & Filename
                If (File.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                    DestFile = AddBkSlash(Destination)
                    ' attempt to copy the file
                    If CopySingleFolder(SourceFile, DestFile, OverWriteFiles) = False Then
                        ' copy failed to complete
                        GoTo Err_CopyFolderList
                    End If
                Else
                    DestFile = AddBkSlash(Destination) & Filename
                    If CopySingleFile(SourceFile, DestFile, OverWriteFiles) = False Then
                        GoTo Err_CopyFolderList
                    End If
                End If
            End If
        Loop While FindNextFile(hSearch, File)
        ' must to do this
        FindClose hSearch
        ' copy was succesful
        CopyFolderList = True
    End If
    
    Exit Function
Err_CopyFolderList:
    If hSearch <> INVALID_HANDLE_VALUE Then
        FindClose hSearch
    End If
    CopyFolderList = False
End Function

' ******************************************************************************
' Routine       : MoveFolder
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/01/01 11:41:45
' Description   : Moves one or more folders from one location to another.
' Inputs        :
'                 source Required.      The path to the folder or folders to be
'                                       moved. The source argument string can
'                                       contain wildcard characters in the last
'                                       path component only.
'                 destination Required. The path where the folder or folders are
'                                       to be moved. The destination argument can't
'                                       contain wildcard characters.
' Outputs       :
' Credits       :
' Modifications :
' Remarks       : If source contains wildcards or
'                 destination ends with a path separator (\),
'                 it is assumed that destination specifies an
'                 existing folder in which to move the matching
'                 files. Otherwise, destination is assumed to
'                 be the name of a destination folder to create.
' ******************************************************************************
Public Function MoveFolder(ByVal Source As String, ByVal Destination As String) As Boolean
    ' if source has wildcards then call the
    ' MoveFolderList() to delete mass files
    ' else simply call MoveSingleFolder() function
    If HasWildcards(Source) Then
        MoveFolder = MoveFolderList(Source, Destination)
    Else
        MoveFolder = MoveSingleFolder(Source, Destination)
    End If
End Function

Private Function MoveSingleFolder(ByVal Source As String, ByVal Destination As String) As Boolean
    Dim lngRetVal As Long
    Dim Folder As String
    
    ' to copy to a folder, destination must end with a path separator
    ' or else we have to create destination
    If Right(Trim(Destination), 1) = "\" Then
        ' strip folder back slash
        Folder = GetFileName(Source)
        Destination = Destination & Folder
    Else
        ' no path separator assume this is a folder to be
        ' created
        CreateFolder Destination
    End If
    lngRetVal = ApiMoveFileEx(Source, Destination, MOVEFILE_REPLACE_EXISTING Or MOVEFILE_COPY_ALLOWED Or MOVEFILE_WRITE_THROUGH)
    MoveSingleFolder = (lngRetVal <> ERROR_SUCCESS)
End Function

Private Function MoveFolderList(ByVal Source As String, ByVal Destination As String) As Boolean
    Dim File As WIN32_FIND_DATA
    Dim SourceFolder As String
    Dim DestFile As String
    Dim Filename As String
    Dim hSearch As Long
    Dim SavPath As String
    
    ' try to find a folder or folders
    hSearch = FindFirstFile(Source, File)
    Source = GetPathName(Source)
    ' check if handle is valid
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do
            ' get the file name
            Filename = TrimNull(File.cFileName)
            If (Filename <> ".") And (Filename <> "..") Then
                ' this must be a directory
                If (File.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                    ' build the folder source path
                    SourceFolder = AddBkSlash(Source) & Filename
                    ' try to move the file(s)
                    If MoveSingleFolder(SourceFolder, Destination) = False Then
                        ' move failed to complete
                        GoTo Err_MoveFolderList
                    End If
                End If
            End If
        Loop While FindNextFile(hSearch, File)
        ' must close search handle
        FindClose hSearch
        ' move succesful
        MoveFolderList = True
    End If
    
    Exit Function
Err_MoveFolderList:
    If hSearch <> INVALID_HANDLE_VALUE Then
        FindClose hSearch
    End If
    MoveFolderList = False
End Function

' ******************************************************************************
' Routine       : DeleteFolder
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/01/01 11:28:46
' Description   : Deletes a specified folder and its contents.
' Inputs        :
'                 folderspec Required.  The name of the folder to delete. The
'                                       folderspec can contain wildcard characters
'                                       in the last path component.
'                 force Optional.       Boolean value that is True if folders with
'                                       the read-only attribute set are to be
'                                       deleted; False (default) if they are not.
' Outputs       : True  - the folder was completely deleted
'                 False - the folder or some folder could not be deleted or were
'                         partially deleted
' Credits       :
' Modifications :
' Remarks       : The DeleteFolder method does not distinguish between folders
'                 that have contents and those that do not. The specified folder
'                 is deleted regardless of whether or not it has contents.
' ******************************************************************************
Public Function DeleteFolder(ByVal FolderSpec As String, Optional Force As Boolean = False) As Boolean
    ' if folderspec has wildcards then call the
    ' DeleteFolderList() to delete mass files and folders
    ' else call DeleteFolderList() function to remove directory
    ' contents then call DeleteSingleFolder to remove folderspec
    ' directory
    If HasWildcards(FolderSpec) Then
        ' remove contents only
        DeleteFolder = DeleteFolderList(FolderSpec, Force)
    Else
        ' remove all directory contents (files and folders)
        If DeleteFolderList(AddBkSlash(FolderSpec) & "*.*", Force) Then
            ' remove current folder
            DeleteFolder = DeleteSingleFolder(FolderSpec, Force)
        End If
    End If
End Function

Private Function DeleteSingleFolder(ByVal FolderSpec As String, Optional Force As Boolean = False) As Boolean
    Dim lngRetVal As Long
    Dim Attr As EFileAttributes
    
    If FolderExists(FolderSpec) Then
        Attr = GetFileAttributes(FolderSpec)
        ' if force was set to true we have to change
        ' folder read-only atribute so it can be
        ' deleted else the function will fail
        If ((Attr And FILE_ATTRIBUTE_READONLY) And (Force = True)) Then
            SetFileAttributes FolderSpec, FILE_ATTRIBUTE_NORMAL
        End If
        lngRetVal = ApiRemoveDirectory(FolderSpec)
        DeleteSingleFolder = (lngRetVal <> ERROR_SUCCESS)
    End If
End Function

Private Function DeleteFolderList(ByVal FolderSpec As String, Optional Force As Boolean = False) As Boolean
    Dim Filename As String
    Dim hSearch As Long
    Dim File As WIN32_FIND_DATA
    Dim Attr As EFileAttributes
    Dim FilePath As String
    
    ' get file or folder attributes
    Attr = GetFileAttributes(FolderSpec)
    ' if is read only but force is true or it is not
    ' read only then continue delete
    If ((Attr And FILE_ATTRIBUTE_READONLY) And Force) Or Not (Attr And FILE_ATTRIBUTE_READONLY) Then
        ' find the first dir item
        hSearch = FindFirstFile(FolderSpec, File)
        FolderSpec = GetPathName(FolderSpec)
        ' che if we have a valid handle to preceed
        If hSearch <> INVALID_HANDLE_VALUE Then
            ' loop directory items to delete
            Do
                ' get the file or folder name
                Filename = TrimNull(File.cFileName)
                ' skip parent and current path
                If (Filename <> ".") And (Filename <> "..") Then
                    ' build the path o be deleted
                    FilePath = AddBkSlash(FolderSpec) & Filename
                    ' check wether it is a file or folder
                    If (File.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                        ' delete the folder recusively
                        If DeleteFolderList(AddBkSlash(FilePath) & "*.*", Force) = False Then
                            GoTo Err_DeleteFolderList
                        End If
                        If DeleteSingleFolder(FilePath, Force) = False Then
                            GoTo Err_DeleteFolderList
                        End If
                    Else
                        ' delete the file
                        If DeleteSingleFile(FilePath, Force) = False Then
                            GoTo Err_DeleteFolderList
                        End If
                    End If
                End If
            Loop While FindNextFile(hSearch, File)
            ' must close search handle
            FindClose hSearch
        End If
        DeleteFolderList = True
    End If
    
    Exit Function
Err_DeleteFolderList:
    ' close search handle
    If hSearch <> INVALID_HANDLE_VALUE Then
        FindClose hSearch
    End If
    ' return false
    DeleteFolderList = False
End Function

' ******************************************************************************
' Routine       : CreateTextFile
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/01/0111:24:02
' Description   : Creates a text file and returns a handle
'                 You must use CloseFile() function to close
'                 files's handle after using the file created
'                 by this function, or else the file will not
'                 grant access while opened by the system
' Inputs        : Filename      The file you wish to create
'                 Overwrite     Overwrite flag which means:
'                     False -   open existing or create a new one
'                     True  -   always create a new file
' Outputs       : handle        File handle
' Credits       :
' Modifications :
' ******************************************************************************
Public Function CreateTextFile(ByVal Filename As String, Optional Overwrite As Boolean = True) As Long
    Dim lngHandle As Long
    Dim FCD As EFileCreationDisposition
    
    ' set the flag creation flag
    If Overwrite Then
        FCD = CREATE_ALWAYS
    Else
        FCD = OPEN_ALWAYS
    End If
    ' open the file to get the filehandle
    lngHandle = ApiCreateFile( _
       Filename, _
       GENERIC_WRITE, _
       FILE_SHARE_READ Or FILE_SHARE_WRITE, _
       ByVal 0&, _
       FCD, _
       FILE_ATTRIBUTE_NORMAL, _
       ByVal 0&)
    CreateTextFile = lngHandle
End Function

' ******************************************************************************
' Routine       : FindFirst
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 12:33:45
' Description   : Replacement for FindFirstFile() Api so that it returns much more
' Inputs        : file information using the TFile structure
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Public Function FindFirst(hSearch As Long, ByVal Path As String, File As TFile) As Boolean
    Dim WFD As WIN32_FIND_DATA
    
    hSearch = FindFirstFile(Path, WFD)
    If hSearch <> INVALID_HANDLE_VALUE Then
        Path = GetPathName(Path)
        GetFileInfo Path, WFD, File
        FindFirst = True
    Else
        FindFirst = False
    End If
End Function

' ******************************************************************************
' Routine       : FindNext
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 04/02/01 11:09:17
' Description   : Find next file or folder pattern
' Inputs        :
' Outputs       :
' Credits       :
' Modifications : Now it closes the search handle when no match is found
' Remarks       :
' ******************************************************************************
Public Function FindNext(hSearch As Long, File As TFile) As Boolean
    Dim WFD As WIN32_FIND_DATA
    Dim lRetval As Long
    
    ' Get next subdirectory.
    lRetval = FindNextFile(hSearch, WFD)
    If lRetval Then
        GetFileInfo GetPathName(File.Path), WFD, File
        FindNext = True
    Else
        FindNext = False
        FindClose hSearch
    End If
End Function

' ******************************************************************************
' Routine       : GetPictureInfo
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 12:34:42
' Description   : Retrieve a BMP, GIF, JPG or PNG image properties
' Inputs        : Filename      name of the file to retrieve image information
'                 PictureInfo   Structure that will contain image properties
' Outputs       :
' Credits       : http://www.4guysfromrolla.com
' Modifications : replaced FileSystemObject for binary file access
' Remarks       :
' ******************************************************************************
Public Function GetPictureInfo(ByVal Filename As String, PictureInfo As TPictureInfo) As Boolean
On Error GoTo Err_GetPictureInfo
    Const constSource As String = m_constClassName & ".GetPictureInfo"

    Dim strPNG As String
    Dim strGIF As String
    Dim strBMP As String
    Dim strType As String
    Dim strBuff As String
    Dim lngSize As Long
    Dim flgFound As Integer
    Dim strTarget As String
    Dim lngPos As Long
    Dim ExitLoop As Boolean
    Dim lngMarkerSize As Long
    Dim TempDepth As String
    Dim hFile As Integer
    
    strType = ""
    GetPictureInfo = False
    strPNG = Chr(137) & Chr(80) & Chr(78)
    strGIF = "GIF"
    strBMP = Chr(66) & Chr(77)

    hFile = FreeFile
    Open Filename For Binary Shared As #hFile
    strType = GetBytes(hFile, 0, 3)
    If strType = strGIF Then               ' is GIF
        PictureInfo.Type = "GIF"
        PictureInfo.Width = lngConvert(GetBytes(hFile, 7, 2))
        PictureInfo.Height = lngConvert(GetBytes(hFile, 9, 2))
        PictureInfo.ColorDepth = 2 ^ ((Asc(GetBytes(hFile, 11, 1)) And 7) + 1)
        GetPictureInfo = True
    ElseIf Left(strType, 2) = strBMP Then      ' is BMP
        PictureInfo.Type = "BMP"
        PictureInfo.Width = lngConvert(GetBytes(hFile, 19, 2))
        PictureInfo.Height = lngConvert(GetBytes(hFile, 23, 2))
        PictureInfo.ColorDepth = 2 ^ (Asc(GetBytes(hFile, 29, 1)))
        GetPictureInfo = True
    ElseIf strType = strPNG Then           ' Is PNG
        PictureInfo.Type = "PNG"
        PictureInfo.Width = lngConvert2(GetBytes(hFile, 19, 2))
        PictureInfo.Height = lngConvert2(GetBytes(hFile, 23, 2))
        TempDepth = GetBytes(hFile, 25, 2)
        Select Case Asc(Right(TempDepth, 1))
            Case 0
                PictureInfo.ColorDepth = 2 ^ (Asc(Left(PictureInfo.ColorDepth, 1)))
                GetPictureInfo = True
            Case 2
                PictureInfo.ColorDepth = 2 ^ (Asc(Left(PictureInfo.ColorDepth, 1)) * 3)
                GetPictureInfo = True
            Case 3
                PictureInfo.ColorDepth = 2 ^ (Asc(Left(PictureInfo.ColorDepth, 1)))  '8
                GetPictureInfo = True
            Case 4
                PictureInfo.ColorDepth = 2 ^ (Asc(Left(PictureInfo.ColorDepth, 1)) * 2)
                GetPictureInfo = True
            Case 6
                PictureInfo.ColorDepth = 2 ^ (Asc(Left(PictureInfo.ColorDepth, 1)) * 4)
                GetPictureInfo = True
            Case Else
                PictureInfo.ColorDepth = -1
        End Select
    Else
        strBuff = GetBytes(hFile, 0, -1)     ' Get all bytes from file
        lngSize = Len(strBuff)
        flgFound = 0
        strTarget = Chr(255) & Chr(216) & Chr(255)
        flgFound = InStr(strBuff, strTarget)
        If flgFound = 0 Then
            Exit Function
        End If
        PictureInfo.Type = "JPG"
        lngPos = flgFound + 2
        ExitLoop = False
        Do While ExitLoop = False And lngPos < lngSize
            Do While Asc(Mid(strBuff, lngPos, 1)) = 255 And lngPos < lngSize
                lngPos = lngPos + 1
            Loop
            If Asc(Mid(strBuff, lngPos, 1)) < 192 Or Asc(Mid(strBuff, lngPos, 1)) > 195 Then
                lngMarkerSize = lngConvert2(Mid(strBuff, lngPos + 1, 2))
                lngPos = lngPos + lngMarkerSize + 1
            Else
                ExitLoop = True
            End If
        Loop
        If ExitLoop = False Then
            PictureInfo.Width = -1
            PictureInfo.Height = -1
            PictureInfo.ColorDepth = -1
        Else
            PictureInfo.Height = lngConvert2(Mid(strBuff, lngPos + 4, 2))
            PictureInfo.Width = lngConvert2(Mid(strBuff, lngPos + 6, 2))
            PictureInfo.ColorDepth = 2 ^ (Asc(Mid(strBuff, lngPos + 8, 1)) * 8)
            GetPictureInfo = True
        End If
    End If
    Close #hFile

Exit Function
Exit_GetPictureInfo:

Exit Function
Err_GetPictureInfo:
    Close #hFile
    GetPictureInfo = False
End Function

Private Function GetBytes(hFile As Integer, Offset As Long, bytes As Long) As String
    Dim lngSize As Long
    Dim strBuff As String
    
    ' First, we get the filesize
    lngSize = LOF(hFile)
    Seek #hFile, 1
    If Offset > 0 Then
        Seek #hFile, Offset
    End If
    If bytes = -1 Then
        strBuff = Space(lngSize - Offset)
    Else
        strBuff = Space(bytes)
    End If
    Get #hFile, , strBuff
    GetBytes = strBuff
End Function

' Functions to convert two bytes to a numeric value (long)
' (both little-endian and big-endian)
Private Function lngConvert(strTemp As String)
    lngConvert = CLng(Asc(Left(strTemp, 1)) + ((Asc(Right(strTemp, 1)) * CLng(256))))
End Function

Private Function lngConvert2(strTemp As String)
    lngConvert2 = CLng(Asc(Right(strTemp, 1)) + ((Asc(Left(strTemp, 1)) * CLng(256))))
End Function

' ******************************************************************************
' Routine       : GetExtensionName
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 04/02/01 11:11:20
' Description   : Returns a string containing the extension name for the last
'                 component in a path.
' Inputs        : path      Required. The path specification for the component
'                           whose extension name is to be returned.
' Outputs       : extension string if any
' Credits       :
' Modifications : Microsoft setup wizard common.bas file
' Remarks       : Changed the name from "Extension()" to "GetExtensionName()"
'                 to maintain module standards
' ******************************************************************************
Public Function GetExtensionName(Path As String) As String
    Dim intDotPos As Integer
    Dim intSepPos As Integer
    Dim FilePath As String
    
    FilePath = Path
    intDotPos = InStrRev(FilePath, ".")
    If intDotPos > 0 Then
        ' we've found a dot. Now make sure there is no '\' after it.
        intSepPos = InStr(intDotPos + 1, FilePath, "\")
        If intSepPos = 0 Then
            ' there is no '\' after the dot. Make sure there is also no '/'.
            intSepPos = InStr(intDotPos + 1, FilePath, "/")
            If intSepPos = 0 Then
                ' the dot has no '\' or '/' after it, so it is good.
                GetExtensionName = UCase(Mid$(FilePath, intDotPos + 1))
            End If
        End If
    End If
End Function

' ******************************************************************************
' Routine       : OpenTextFile
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 04/02/01 11:16:05
' Description   : Opens a specified file and returns a handle that can be used
'                 to read from or append to the file.
' Inputs        : filename Required. String expression that identifies the file
'                                    to open.
'                 iomode Optional.   Indicates input/output mode. Can be one of
'                                    two constants, either ForReading or ForAppending.
'                 create Optional.   Boolean value that indicates whether a new file
'                                    can be created if the specified filename doesn't
'                                    exist. The value is True if a new file is created;
'                                    False if it isn't created. The default is False.
' Outputs       : Handle for the opened file
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Public Function OpenTextFile(Filename As String, Optional IOMode As EIOModes = ForReading, Optional Create As Boolean = False) As Long
    On Error GoTo Err_OpenTextFile
    Const constSource As String = m_constClassName & ".OpenTextFile"
   
    Dim hFile As Integer
    Dim Access As Long
    Dim CD As EFileCreationDisposition
    Dim ShMode As Long
   
    ShMode = FILE_SHARE_WRITE Or FILE_SHARE_READ
    
    If IOMode And ForAppending Then
        Access = Access Or GENERIC_WRITE
    Else
        If IOMode And ForReading Then
            Access = Access Or GENERIC_READ
        End If
        If IOMode And ForWriting Then
            Access = Access Or GENERIC_WRITE
        End If
    End If
    If Create = True Then
        CD = OPEN_ALWAYS
    Else
        CD = OPEN_EXISTING
    End If
    ' open the file
    hFile = ApiCreateFile(Filename, Access, ShMode, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    ' check for appending mode
    ' we must go to the bottom of the file
    If hFile <> INVALID_HANDLE_VALUE Then
        If IOMode And ForAppending Then
            Seek #hFile, LOF(hFile) + 1
        End If
    End If
    OpenTextFile = hFile
    Exit Function
   
    Exit Function
Err_OpenTextFile:
    OpenTextFile = INVALID_HANDLE_VALUE
End Function

' ******************************************************************************
' Routine       : CloseFile
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 04/02/01 11:21:21
' Description   : Flush file buffer and closes the file handle
' Inputs        : hFile Required. File handle to close
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Public Function CloseFile(ByVal hFile As Long) As Boolean
    Dim lngRetVal As Long
        
    ' flush the file buffers to force writing of the data.
    lngRetVal = FlushFileBuffers(hFile)
    ' close file handle
    lngRetVal = ApiCloseHandle(hFile)
    ' return success
    CloseFile = (lngRetVal <> ERROR_SUCCESS)
End Function

' *******************************************************
' Routine Name : (PUBLIC in MODULE) Function ReturnApiErrString
' Written By   : L.J. Johnson
' Programmer   : L.J. Johnson [Slightly Tilted Software]
' Date Writen  : 01/16/1999 -- 12:56:46
' Inputs       : ErrorCode:Long - Number returned from API error
' Outputs      : N/A
' Description  : Function returns the error string
'              : The original code appeared in Keith Pleas
'              :     article in VBPJ, April 1996 (OLE Expert
'              :     column).  Thanks, Keith.
' *******************************************************
Private Function ReturnApiErrString(ErrorCode As Long) As String
    On Error Resume Next ' Don't accept an error here
    Dim strBuffer As String

    ' allocate the string, then get the system
    ' to tell us the error message associated
    ' with this error number
    strBuffer = String(256, 0)
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM _
       Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, _
       ErrorCode, 0&, strBuffer, Len(strBuffer), 0&
    ' strip the last null, then the last CrLf
    ' pair if it exists
    strBuffer = Left(strBuffer, InStr(strBuffer, vbNullChar) - 1)
    If Right$(strBuffer, 2) = Chr$(13) & Chr$(10) Then
        strBuffer = Mid$(strBuffer, 1, Len(strBuffer) - 2)
    End If
    ' set the return value
    ReturnApiErrString = strBuffer
End Function

' ******************************************************************************
' Routine       : GetFileContentType
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 12:58:30
' Description   : Retrieve the content-type of a file from registry
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Remarks       : Files content type are locate at
'                 HKEY_CLASSES_ROOT\.[Extension]\Content Type
'
' ******************************************************************************
Public Function GetFileContentType(ByVal Filename As String) As String
    Dim hKey As Long
    Dim SubKey As String
    Dim Ext As String
    
    ' extract file extension
    Ext = "." & GetExtensionName(Filename)
    hKey = HKEY_CLASSES_ROOT
    SubKey = Ext
    GetFileContentType = GetRegValue(hKey, SubKey, "Content Type", REG_SZ, vbNullString)
End Function

' ******************************************************************************
' Routine       : BuildPath
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 04/02/01 11:22:33
' Description   : Appends a name to an existing path.
' Inputs        : path Required. Existing path to which name is appended. Path can
'                                be absolute or relative and need not specify an
'                                existing folder.
'                 name Required. Name being appended to the existing path.

' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Public Function BuildPath(Path As String, Name As String) As String
    Dim FilePath As String
    
    FilePath = Path
    FilePath = ConcatString(FilePath, Name, "\")
    BuildPath = FilePath
End Function

' ******************************************************************************
' Routine       : GetFullPathName
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 13:29:01
' Description   : Retrieve fully qualified path/name specs.
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Public Function GetFullPathName(Path As String) As String
    Dim Buffer As String
    Dim nFilePart As Long
    Dim nRet As Long
   
    Buffer = Space(MAX_PATH)
    nRet = ApiGetFullPathName(Path, Len(Buffer), Buffer, nFilePart)
    If nRet Then
        GetFullPathName = Left(Buffer, nRet)
    End If
End Function

' ******************************************************************************
' Routine       : IsReadOnly
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 04/02/01 11:23:46
' Description   : check if file or folder is read only
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Public Function IsReadOnly(FilePath As String) As Boolean
    Dim Attr As EFileAttributes
   
    Attr = GetFileAttributes(FilePath)
    IsReadOnly = Attr And FILE_ATTRIBUTE_READONLY
End Function

' ******************************************************************************
' Routine       : IsHidden
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 04/02/01 11:24:10
' Description   : check if file or folder is hidden
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Public Function IsHidden(FilePath As String) As Boolean
    Dim Attr As EFileAttributes
   
    Attr = GetFileAttributes(FilePath)
    IsHidden = Attr And FILE_ATTRIBUTE_HIDDEN
End Function

' ******************************************************************************
' Routine       : IsSystem
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 04/02/01 11:24:27
' Description   : check if file or folder is system
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Public Function IsSystem(FilePath As String) As Boolean
    Dim Attr As EFileAttributes
   
    Attr = GetFileAttributes(FilePath)
    IsSystem = Attr And FILE_ATTRIBUTE_SYSTEM
End Function

' ******************************************************************************
' Routine       : IsArchive
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 04/02/01 11:24:41
' Description   : check if file or folder isan archive
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Public Function IsArchive(FilePath As String) As Boolean
    Dim Attr As EFileAttributes
   
    Attr = GetFileAttributes(FilePath)
    IsArchive = Attr And FILE_ATTRIBUTE_ARCHIVE
End Function

' ******************************************************************************
' Routine       : IsTemporary
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 04/02/01 11:24:49
' Description   : check if file or folder is temporary
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Public Function IsTemporary(FilePath As String) As Boolean
    Dim Attr As EFileAttributes
   
    Attr = GetFileAttributes(FilePath)
    IsTemporary = Attr And FILE_ATTRIBUTE_TEMPORARY
End Function

' ******************************************************************************
' Routine       : IsCompressed
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 04/02/01 11:25:04
' Description   : check if file or folder iscompressed
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Public Function IsCompressed(FilePath As String) As Boolean
    Dim Attr As EFileAttributes
   
    Attr = GetFileAttributes(FilePath)
    IsCompressed = Attr And FILE_ATTRIBUTE_COMPRESSED
End Function

' ******************************************************************************
' Routine       : IsDirectory
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 04/02/01 11:25:13
' Description   : check if file or folder isa directory
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Public Function IsDirectory(FilePath As String) As Boolean
    Dim Attr As EFileAttributes
   
    Attr = GetFileAttributes(FilePath)
    IsDirectory = Attr And FILE_ATTRIBUTE_DIRECTORY
End Function

' ******************************************************************************
' Routine       : GetShortPathName
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 04/02/01 11:25:25
' Description   : get path short name
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Public Function GetShortPathName(Path As String) As String
    Dim Buffer As String
    Dim nRet As Long
    
    ' retrieve short path name.
    Buffer = Space(MAX_PATH)
    nRet = ApiGetShortPathName(Path, Buffer, Len(Buffer))
    If nRet Then
        GetShortPathName = Left(Buffer, nRet)
        'm_PathShort = Left(m_PathNameShort, Len(m_PathNameShort) - Len(m_NameShort))
    End If
End Function

' ******************************************************************************
' Routine       : GetTempPath
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 14:07:05
' Description   : Returns a randomly generated temporary folder name that is
'                 useful for performing operations that require a temporary folder.
' Inputs        :
' Outputs       :
' Credits       : I really do not remember.
' Modifications :
' Remarks       :
' ******************************************************************************
Private Function GetTempPath() As String
    Dim r As Long
    Dim sWinTmpDir As String
   
    'get the user's windows\temp folder
    'pad the passed string
    sWinTmpDir = Space$(MAX_PATH)
   
    'get the folder
    r = ApiGetTempPath(MAX_PATH, sWinTmpDir)
    'r contains the number of chrs up to the
    'terminating null, so a simple left$ can
    'be used. Its also conveniently terminated
    'with a slash.
    sWinTmpDir = Left$(sWinTmpDir, r)
    ' return path
    GetTempPath = sWinTmpDir
End Function

' ******************************************************************************
' Routine       : GetTempFileName
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 14:06:39
' Description   : Returns a randomly generated temporary file name that
'                 is useful for performing operations that require a temporary file
' Inputs        :
' Outputs       :
' Credits       : I really do not remember
' Modifications :
' Remarks       :
' ******************************************************************************
Public Function GetTempFileName(Optional sWinTmpDir As String, Optional Prefix As String = "Tmp") As String
    Dim r As Long
    Dim sTmpFile As String
   
    ' pad a working string
    sTmpFile = Space$(MAX_PATH)
    
    If sWinTmpDir = "" Then sWinTmpDir = GetTempPath
    
    'call the API.
    'The first param is the path in which to create
    'the file. Passing "." creates the file in the
    'current directory. A specific path can also be
    'passed. Using the GetTempPath() API (as in
    'Form Load) returns Windows temporary folder,
    'which is used here.
  
    'the second param is the prefix string (note:
    'null-terminated under NT). The function uses
    'up to the first three characters of this string
    'as the prefix of the filename. This string must
    'consist of characters in the ANSI character set.
  
    'the third parameter, uUnique, specifies an
    'unsigned integer that the function converts to
    'a hexadecimal string for use in creating the
    'temporary filename. If uUnique is nonzero, the
    'function appends the hexadecimal string to
    'lpPrefixString to form the temporary filename.
    'In this case, the function does not create the
    'specified file, and does not test whether the
    'filename is unique.

    'If uUnique is zero, as below, the function uses
    'a hexadecimal string derived from the current
    'system time. In this case, the function uses
    'different values until it finds a unique filename,
    'and then it creates the file in the lpPathName directory.
  
    'the last param is the variable to contain the temporary
    'filename, null-terminated and consisting of characters
    'in the ANSI character set. This string should be padded
    'at least the length, in bytes, specified by MAX_PATH to
    'accommodate the path.
    r = ApiGetTempFileName(sWinTmpDir, Prefix, 0, sTmpFile)
    'If the function succeeds, the return value 'r'
    'specifies the unique numeric value (in decimal)
    'used in the temporary filename.
    '
    'If the function fails, the return value is zero.
    If r <> 0 Then
        'strip the trailing null
        sTmpFile = Left$(sTmpFile, InStr(sTmpFile, Chr$(0)) - 1)
        'display the created file, the decimal
        'value and the hex value in the list
        GetTempFileName = sTmpFile
    End If
End Function

' ******************************************************************************
' Routine       : GetSpecialFolder
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 14:30:06
' Description   : Returns the special folder specified
' Inputs        : specialfolder:     Required. The name of the special folder to be returned. Can be any of the constants shown in the Settings section.
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Public Function GetSpecialFolder(SpecialFolder As ESpecialFolders) As String
    Dim Bufstr As String
    Dim nRet As Long
    
    Bufstr = Space$(MAX_PATH)
    If SpecialFolder = WindowsFolder Then
        nRet = ApiWindDir(Bufstr, MAX_PATH)
    ElseIf SpecialFolder = SystemFolder Then
        nRet = ApiSysDir(Bufstr, MAX_PATH)
    ElseIf SpecialFolder = TemporaryFolder Then
        nRet = ApiGetTempPath(MAX_PATH, Bufstr)
    End If
    If nRet > 0 Then
        GetSpecialFolder = Left(Bufstr, nRet)
    End If
End Function

' ******************************************************************************
' Routine       : FormatFileSize
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 15:23:38
' Description   : Format a file size to KB, MB, GB or bytes
' Inputs        :
' Outputs       :
' Credits       : Karl E. Peterson (http://www.mvps.org/vb)
' Modifications :
' Remarks       :
' ******************************************************************************
Public Function FormatFileSize(ByVal Size As Long) As String
    Dim sRet As String
    Const KB& = 1024
    Const MB& = KB * KB
   
    ' Return size of file in kilobytes.
    If Size < KB Then
        sRet = Format(Size, "#,##0") & " bytes"
    Else
        Select Case Size \ KB
            Case Is < 10
                sRet = Format(Size / KB, "0.00") & "KB"
            Case Is < 100
                sRet = Format(Size / KB, "0.0") & "KB"
            Case Is < 1000
                sRet = Format(Size / KB, "0") & "KB"
            Case Is < 10000
                sRet = Format(Size / MB, "0.00") & "MB"
            Case Is < 100000
                sRet = Format(Size / MB, "0.0") & "MB"
            Case Is < 1000000
                sRet = Format(Size / MB, "0") & "MB"
            Case Is < 10000000
                sRet = Format(Size / MB / KB, "0.00") & "GB"
        End Select
        sRet = sRet & " (" & Format(Size, "#,##0") & " bytes)"
    End If
    FormatFileSize = sRet
End Function

' ******************************************************************************
' Routine       : GetFile
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 15:58:55
' Description   : Get a file an populate File structure with file extended info
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Public Function GetFile(FilePath As String, File As TFile) As Boolean
    Dim hSearch As Long
    Dim WFD As WIN32_FIND_DATA
    
    hSearch = FindFirstFile(FilePath, WFD)
    If hSearch <> INVALID_HANDLE_VALUE Then
        GetFileInfo GetPathName(FilePath), WFD, File
        FindClose hSearch
        GetFile = True
    Else
        GetFile = False
    End If
End Function

' ******************************************************************************
' Routine       : FileInUse
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 15:58:19
' Description   : Check if a file is in use by another process
' Inputs        :
' Outputs       :
' Credits       : Microsoft Setup wizard (common.bas)
' Modifications :
' Remarks       :
' ******************************************************************************
Public Function FileInUse(FileSpec As String) As Boolean
    Dim hFile As Long
    Dim strPathName As String
    
    strPathName = FileSpec
    On Error Resume Next
    ' Remove any trailing directory separator character
    If Right$(strPathName, 1) = "\" Then
        strPathName = Left$(strPathName, Len(strPathName) - 1)
    End If
    hFile = ApiCreateFile(strPathName, GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL Or FILE_FLAG_WRITE_THROUGH, 0)
    If hFile = INVALID_HANDLE_VALUE Then
        FileInUse = Err.LastDllError = ERROR_SHARING_VIOLATION
    Else
        ApiCloseHandle hFile
    End If
    Err.Clear
End Function

' ******************************************************************************
' Routine       : GetTimeZone
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 16:06:24
' Description   : get time zone information
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Private Function GetTimeZone() As Long
    Dim objTimeZone As TIME_ZONE_INFORMATION
    Dim lngResult As Long
    Dim i As Long
    
    lngResult = GetTimeZoneInformation&(objTimeZone)
    Select Case lngResult
        Case 0&, 1& 'use standard time
            GetTimeZone = -(objTimeZone.Bias + objTimeZone.StandardBias) 'into minutes
        Case 2& 'use daylight savings time
            GetTimeZone = -(objTimeZone.Bias + objTimeZone.DaylightBias) 'into minutes
    End Select
End Function

' ******************************************************************************
' Routine       : LocalDateToUTC
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 16:06:39
' Description   : Convert local date to UTC format "Sat, 23 Jan 2001 10:00:00 GMT"
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Public Function LocalDateToUTC(ByVal LocalDate As Date) As String
    Dim lngTZ As Long
    Dim D As Date
    Dim strTemp As String
    
    lngTZ = GetTimeZone
    D = DateSerial(Year(LocalDate), Month(LocalDate), Day(LocalDate)) + TimeSerial(Hour(LocalDate), Minute(LocalDate) - lngTZ, Second(LocalDate))
    LocalDateToUTC = Format(D, "ddd, d mmm yyyy hh:mm:ss") & " GMT"
End Function

' ******************************************************************************
' Routine       : UTCToLocalDate
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 16:07:30
' Description   : Convert UTC date to local user date "mm/dd/yyyy hh:mm:ss"
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Public Function UTCToLocalDate(ByVal GMTDate As String) As Variant
    Dim dt As Date
    Dim D As Date
    Dim lngTZ As Long
    Dim aData As Variant
    
    '--> Wed, 26 dez 2000 02:00:00 GMT
    On Error Resume Next
    lngTZ = GetTimeZone
    aData = Split(GMTDate, " ")
    If UBound(aData) <> 5 Then
        UTCToLocalDate = Null
        Exit Function
    End If
    dt = DateSerial(aData(3), Mtoi(aData(2)), aData(1)) & " " & aData(4)
    D = DateSerial(Year(dt), Month(dt), Day(dt)) + TimeSerial(Hour(dt), Minute(dt) + lngTZ, Second(dt))
    UTCToLocalDate = D
End Function

' ******************************************************************************
' Routine       : Mtoi
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 29/01/01 16:08:30
' Description   : Convert an abreviated month name to an integer
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Private Function Mtoi(m As Variant) As Integer
    Dim i As Integer
    
    For i = 1 To 12
        If MonthName(i, True) = Left(m, 3) Then
            Mtoi = i
            Exit Function
        End If
    Next
    Mtoi = 1
End Function

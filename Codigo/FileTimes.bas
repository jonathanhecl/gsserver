Attribute VB_Name = "FileTimes"
'**************************************
'Windows API/Global Declarations for :Fi
'     leTimes
'**************************************


Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long


Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Public Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long


Public Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long


Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
    Public Const OFS_MAXPATHNAME = 128
    Public Const GENERIC_WRITE = &H40000000
    Public Const GENERIC_READ = &H80000000


Public Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
    End Type


Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
    End Type


Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
    End Type


Public Function GetFileInfo(strFile As String) As Date
    Dim FTimeCreated As FILETIME
    Dim FTimeModified As FILETIME
    Dim FTimeAccessed As FILETIME
    Dim FLocalTime As FILETIME
    Dim lngReturn As Long
    Dim lngFHandle As Long
    Dim OpenBuffer As OFSTRUCT
    Dim TSystem As SYSTEMTIME
    Dim strResponse As String
    'Open the file up, and get the handle
    lngFHandle = OpenFile(strFile, OpenBuffer, GENERIC_READ)
    'Use the handle to get the created, open
    '     ed and modified times
    lngReturn = GetFileTime(lngFHandle, FTimeCreated, FTimeAccessed, FTimeModified)
    'Close the file
    CloseHandle lngFHandle
    'Get the file size
    'strResponse = "File: " & strFile & ", size: " & FileLen(strFile) & " bytes." & vbCrLf
    'Convert the created time to local file
    '     time
    'Call FileTimeToLocalFileTime(FTimeCreated, FLocalTime)
    'Convert the local file time to system t
    '     ime
    'Call FileTimeToSystemTime(FLocalTime, TSystem)
    'Write the response
    'strResponse = strResponse & "Created on: " & TSystem.wYear & "-" & TSystem.wMonth & "-" & TSystem.wDay & " at " & TSystem.wHour & ":" & TSystem.wMinute & ":" & TSystem.wSecond & vbCrLf
    'Do again for other times (modified and
    '     accessed)
    Call FileTimeToLocalFileTime(FTimeModified, FLocalTime)
    Call FileTimeToSystemTime(FLocalTime, TSystem)
    GetFileInfo = TSystem.wDay & "/" & TSystem.wMonth & "/" & TSystem.wYear
    'strResponse = TSystem.wYear & "-" & TSystem.wMonth & "-" & TSystem.wDay   '& " at " & TSystem.wHour & ":" & TSystem.wMinute & ":" & TSystem.wSecond & vbCrLf
    'Call FileTimeToLocalFileTime(FTimeAccessed, FLocalTime)
    'Call FileTimeToSystemTime(FLocalTime, TSystem)
    'strResponse = strResponse & "Accessed on: " & TSystem.wYear & "-" & TSystem.wMonth & "-" & TSystem.wDay & " at " & TSystem.wHour & ":" & TSystem.wMinute & ":" & TSystem.wSecond & vbCrLf
    'Return the response
    'GetFileInfo = strResponse
End Function

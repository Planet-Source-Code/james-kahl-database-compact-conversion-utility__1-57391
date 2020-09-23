Attribute VB_Name = "mMain"
'****************************************************************************************
'Module:        mMain - BAS Module
'Filename:      mMain.bas
'Author:        Jim Kahl
'Purpose:       Sub Main when acting as standalone, other general routines
'****************************************************************************************
Option Explicit

'****************************************************************************************
'API CONSTANTS
'****************************************************************************************
Private Const INVALID_HANDLE_VALUE As Long = (-1)
Private Const MAX_PATH As Long = 260
Private Const ICC_USEREX_CLASSES = &H200

'****************************************************************************************
'API TYPES
'****************************************************************************************
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
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

Private Type INITCOMMONCONTROLSTYPE
   dwSize As Long
   dwICC As Long
End Type

'****************************************************************************************
'API FUNCTIONS
'****************************************************************************************
Private Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32.dll" _
        Alias "FindFirstFileA" ( _
                ByVal lpFileName As String, _
                lpFindFileData As WIN32_FIND_DATA) _
                As Long
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" ( _
                iccex As INITCOMMONCONTROLSTYPE) _
                As Boolean

'****************************************************************************************
'METHODS - PUBLIC
'****************************************************************************************
Public Sub Main()
    'Purpose:       controlling routine for the application
    On Error Resume Next
   
    Dim uICC As INITCOMMONCONTROLSTYPE
   
    'show the standard vb controls as XP themes
    'NOTE: there needs to be an appropriately named .exe.manifest file in the
    'App.Path in order for this to work.
    With uICC
        .dwSize = LenB(uICC)
        .dwICC = ICC_USEREX_CLASSES
    End With
    InitCommonControlsEx uICC
    
    'show the main form
    frmMain.Show vbModal
    
End Sub

Public Function FileExists(ByRef PathName As String) As Boolean
    'Purpose:       determines if a file exists
    'Parameters:    PathName - a fully qualified path and filename
    'Returns:       True - the file exists
    '               False - the file either does not exist, or can not be found at the
    '                   path passed in the PathName parameter
    'Assumes:       PathName must be a valid drive:\path\filename.ext format or
    '                   \\server\share\path\filename.ext format
    Dim tFD As WIN32_FIND_DATA
    Dim hFile As Long
    
    hFile = FindFirstFile(PathName, tFD)
    If hFile <> INVALID_HANDLE_VALUE Then
        FileExists = True
        Call FindClose(hFile)
    End If
End Function

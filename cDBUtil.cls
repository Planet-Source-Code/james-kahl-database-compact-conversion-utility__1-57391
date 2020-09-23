VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDBUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************
'Module:        cDBUtil - Class Module
'Filename:      cDBUtil.cls
'Author:        Jim Kahl
'Purpose:       mainly routines that will help administrators and developers maintain
'               the databases necessary for operation of this application
'               also used for reading table names in the limits database that are used to
'               define the steps/measurements needed for execution of the test.
'Dependencies:  Microsoft Access 10.0 Object Library - msacc.olb
'****************************************************************************************
Option Explicit

'****************************************************************************************
'VARIABLES - PRIVATE
'****************************************************************************************
Private msDBName As String

'****************************************************************************************
'PROPERTIES - PUBLIC READ/WRITE
'****************************************************************************************
Public Property Get DBPath() As String
    DBPath = msDBName
End Property

Public Property Let DBPath(ByVal Filename As String)
    msDBName = Filename
End Property

'****************************************************************************************
'METHODS - PUBLIC
'****************************************************************************************
Public Function Archive() As Boolean
    'TODO: Write procedure to archive database
    'this can be either to archive only specific records - such as by dates
    'or it can be to archive the entire database and then start with a fresh
    'blank copy
    MsgBox "Option Not Available", vbInformation, App.Title
End Function

Public Function Backup() As Boolean
    'TODO: write procedure to create a backup of the database
    'need to decide how we want to handle backups - ie. same name with date in filename
    'or whether we want to give user the option of selecting what name to give the
    'backup by means of a dialog
    'NOTE: There is a Backup routine in CustBase.vbp
    MsgBox "Option Not Available", vbInformation, App.Title
End Function

Public Function Convert( _
        Optional ByVal Format As AcFileFormat = acFileFormatAccess2002) As Boolean
    'Purpose:       provide a quick and easy way to convert a database from one version
    '               of Access to another without having to open Access
    'Parameters:    Format - an enumerated constant or value that equates to the version
    '                   of access we want to convert to
    'Returns:       True - the function was successful
    '               False - the function failed
    
    Dim sPath As String
    Dim oAccess As New Access.Application
    
    Const EXT As String = ".TMP"
    
    On Error GoTo ErrHandler
    
    Screen.MousePointer = vbHourglass
    
    'The temp file is put in the same directory as the database, and is named with the
    'database name, plus the EXT TMP. The temp file is deleted when the proc is done.
    sPath = getPathFromFullPath(msDBName) & getNamePartNoExt(msDBName) & EXT
    
    'make sure that any old temp database doesn't exist
    If FileExists(sPath) Then
        Kill sPath
    End If
    
    'convert the database to the format desired
    oAccess.ConvertAccessProject msDBName, sPath, Format
    
    'destroy the instance
    Set oAccess = Nothing
    
    'delete the old file
    Kill msDBName
    
    'copy the file to the old filename
    FileCopy sPath, msDBName
    
    'kill the temp file
    Kill sPath
    
    MsgBox "The Database: " & msDBName & " was successfully ocnverted.", vbInformation, _
            App.Title
    
    Convert = True
    
ExitProc:
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrHandler:
    Dim lRet As VbMsgBoxResult
    Dim sMsg As String
    
    sMsg = Err.Number & ": " & Err.Description
    
    lRet = MsgBox(sMsg & vbCrLf & vbCrLf & "Would you like to try again?", vbYesNo, _
            App.Title)
    
    If lRet = vbYes Then
        Compact
    Else
        Resume ExitProc
    End If
        
End Function

Public Function Restore() As Boolean
    'TODO: write procedure to restore a backed up database
    'this will probably need to have some type of dialog associated with it
    'to give user the choice of which database to restore from
    'NOTE: there is a Restore routine in CustBase.vbp
    MsgBox "Option Not Available", vbInformation, App.Title
End Function

Public Function Compact() As Long
    'Purpose:       this routine will bypass MS Access and use the JET Engine directly to
    '               compact and repair the database that is pointed to by the Database
    '               property of this class object to a temporary file, it will then
    '               delete the old database file and rename the new compacted file to the
    '               old database filename
    'Returns:       If Successful the function will return the length of the new database
    '                   file
    '               If the function fails it will return an error code defined by the JET
    '                   Engine
    'Assumes:       the Database property must hold a valid file spec to an MS Access
    '                   database
    Dim sPath As String
    Dim oAccess As New Access.Application
    
    Const EXT As String = ".TMP"
    
    'use jet engine to compact a database
    'this should be done at periodic intervals for the databases that are
    'written to regularly, could be done automatically or by means of
    'a dialog
    On Error GoTo ErrHandler
    
    'The temp file is put in the same directory as the database, and is named with the
    'database name, plus the EXT TMP. The temp file is deleted when the proc is done.
    sPath = getPathFromFullPath(msDBName) & getNamePartNoExt(msDBName) & EXT
    
    'make sure that any old compacted database doesn't exist
    If FileExists(sPath) Then
        Kill sPath
    End If
    
    'compact the database
    oAccess.CompactRepair msDBName, sPath
    Set oAccess = Nothing
    
    'delete the original file, then copy the temp file to the original name,
    'and then delete the temp file.
    Kill msDBName
    
    FileCopy sPath, msDBName
    
    Kill sPath
    
    'return the file size of the new filename
    Compact = FileLen(msDBName)
    
    MsgBox "The Database: " & msDBName & " was successfully compacted.", vbInformation, _
            App.Title
    Exit Function
    
ErrHandler:
    Dim lRet As VbMsgBoxResult
    Dim sMsg As String
    
    sMsg = Err.Number & ": " & Err.Description
    
    lRet = MsgBox(sMsg & vbCrLf & vbCrLf & "Would you like to try again?", vbYesNo, _
            App.Title)
    
    If lRet = vbYes Then
        Compact
    End If
        
End Function

'****************************************************************************************
'METHODS - PRIVATE
'****************************************************************************************
Private Function getNamePartNoExt(ByRef Path As String) As String
    'Purpose:       to get the filename with no extension
    'Parameters:    Path - a reference to a file spec to test
    'Returns:       If the function is successful the return value will be the path spec
    '               and filename of the file in the path without the extension of the
    '               file
    'Assumes:       Path must be a valid Path and Filename which can be either format of
    '               drive:\path\file.extension or \\server\sharename\path\file.extension
    Dim lPos As Long
    Dim sTemp As String
    
    lPos = InStrRev(Path, "\")
    
    sTemp = Right$(Path, Len(Path) - lPos)

    lPos = InStr(sTemp, ".")
    
    sTemp = Left$(sTemp, lPos - 1)
    
    getNamePartNoExt = sTemp
    
End Function

Private Function getPathFromFullPath(ByRef Path As String) As String
    'Parameters:    Path - a reference to a file spec to test
    'Returns:       If the function is successful the return value will be the path spec
    '               of the file in the path without the filename and extension of the
    '               file
    'Assumes:       Path must be a valid Path and Filename which can be either format of
    '               drive:\path\file.extension or \\server\sharename\path\file.extension
    Dim lPos As Long
    Dim sTemp
    
    lPos = InStrRev(Path, "\")
    
    sTemp = Left$(Path, lPos)
    
    getPathFromFullPath = sTemp
    
End Function

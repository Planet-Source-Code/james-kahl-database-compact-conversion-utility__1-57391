VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInifile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************
'Module:        cIniFile - Class Module
'Filename:      cIniFile.cls
'Original:      obtained from vbAcclerator.com
'Copyright:     Steve McMahon 1997 - see copyright info below
'Modified:      Jim Kahl
'Modifications:
'   1 - changed variable names and routines to match PGA-ETS naming conventions
'   2 - modified white space for readability
'   3 - added this comment header
'   4 - added other comments as deemed necessary
'   5 - explicitly declared parameters as being passed ByVal or ByRef depending on usage
'   6 - rearranged source to group like items of code together
'   7 - changed some variable types since some were strings or integers and should have
'       been Longs
'   8 - explicitly declared properties and procedures as public for clarity
'NOTE:          strings in this class have not been added to the resource file so this
'               class can be a direct drop-in class
'****************************************************************************************
' =========================================================
' Class:    cIniFile
' Author:   Steve McMahon
' Date  :   21 Feb 1997
'
' A nice class wrapper around the INIFile functions
' Allows searching,deletion,modification and addition
' of Keys or Values.
'
' Updated 10 May 1998 for VB5.
'   * Added EnumerateAllSections method
'   * Added Load and Save form position methods
' =========================================================
Option Explicit

'****************************************************************************************
'API FUNCTIONS
'****************************************************************************************
Private Declare Function WritePrivateProfileString Lib "KERNEL32" _
        Alias "WritePrivateProfileStringA" ( _
                ByVal lpApplicationName As String, _
                ByVal lpKeyName As Any, _
                ByVal lpString As Any, _
                ByVal lpFileName As String) _
                As Long
Private Declare Function GetPrivateProfileString Lib "KERNEL32" _
        Alias "GetPrivateProfileStringA" ( _
                ByVal lpApplicationName As Any, _
                ByVal lpKeyName As Any, _
                ByVal lpDefault As Any, _
                ByVal lpReturnedString As String, _
                ByVal nSize As Long, _
                ByVal lpFileName As String) _
                As Long


'****************************************************************************************
'VARIABLES - PRIVATE
'****************************************************************************************
Private msPath As String
Private msKey As String
Private msSection As String
Private msDefault As String
Private mlLastRetCode As Long

'****************************************************************************************
'PROPERTIES - PUBLIC READ/WRITE
'****************************************************************************************
'Default - sets/returns the default value for a given key in the ini file
Public Property Let Default(ByRef sDefault As String)
    msDefault = sDefault
End Property

Public Property Get Default() As String
    Default = msDefault
End Property
'*******************************************
'Path - sets/returns the path and filename for the file we want to utilize
Public Property Let Path(ByRef sPath As String)
    msPath = sPath
End Property

Public Property Get Path() As String
    Path = msPath
End Property
'*******************************************
'Key - sets/returns the item that we want to get/set information for
Public Property Let Key(ByRef sKey As String)
    msKey = sKey
End Property

Public Property Get Key() As String
    Key = msKey
End Property
'*******************************************
'Section - sets/returns the group of items that a specific item belongs to
Public Property Let Section(sSection As String)
    msSection = sSection
End Property

Public Property Get Section() As String
    Section = msSection
End Property
'*******************************************
'Value - sets/returns the value of a specific item as pointed to by the Path, Section
'   and Key properties
Public Property Get Value() As String
    Dim sBuf As String
    Dim lSize As Long
    Dim lRet As Long

    'setup the buffer and size of the buffer to pass to the api
    sBuf = Space$(255)
    lSize = Len(sBuf)
    
    lRet = GetPrivateProfileString(msSection, msKey, msDefault, sBuf, lSize, msPath)
    
    If (lSize > 0) Then
        Value = Left$(sBuf, lRet)
    Else
        Value = vbNullString
    End If

End Property

Public Property Let Value(ByRef sValue As String)
    Dim iPos As Integer
    
    ' Strip chr$(0):
    iPos = InStr(sValue, Chr$(0))
    
    Do While iPos <> 0
        sValue = Left$(sValue, (iPos - 1)) & Mid$(sValue, (iPos + 1))
        iPos = InStr(sValue, Chr$(0))
    Loop
    
    mlLastRetCode = WritePrivateProfileString(msSection, msKey, sValue, msPath)
End Property
'*******************************************
'INISection - sets/returns the name of a section in the file
Property Get INISection() As String
    Dim sBuf As String
    Dim lSize As Long
    Dim lRet As Long

    'set up the buffer and size of the buffer to send to the API
    sBuf = Space$(8192)
    lSize = Len(sBuf)
    
    lRet = GetPrivateProfileString(msSection, 0&, msDefault, sBuf, lSize, msPath)
    
    If (lSize > 0) Then
        INISection = Left$(sBuf, lRet)
    Else
        INISection = vbNullString
    End If

End Property

Property Let INISection(ByRef sSection As String)
    mlLastRetCode = WritePrivateProfileString(msSection, 0&, sSection, msPath)
End Property

'****************************************************************************************
'PROPERTIES - PUBLIC READ ONLY
'****************************************************************************************
Public Property Get LastReturnCode() As Long
    'used to determine what the return value from an API was
    LastReturnCode = mlLastRetCode
End Property

Public Property Get Success() As Boolean
    Success = (mlLastRetCode <> 0)
End Property

Public Property Get Sections() As String
    'Purpose:       to provide the developer with a way to determine and enumerate the
    '               number of sections contained in the INI File
    Dim sBuf As String
    Dim lSize As Long
    Dim lRet As Long

    sBuf = Space$(8192)
    lSize = Len(sBuf)
    
    'set up the buffer and size of the buffer to send to the API
    lRet = GetPrivateProfileString(0&, 0&, msDefault, sBuf, lSize, msPath)
    
    If (lSize > 0) Then
        Sections = Left$(sBuf, lRet)
    Else
        Sections = vbNullString
    End If

End Property

'****************************************************************************************
'METHODS - PUBLIC
'****************************************************************************************
Public Sub DeleteKey()
    'Purpose:       to delete an item from the given section
    mlLastRetCode = WritePrivateProfileString(msSection, msKey, 0&, msPath)
End Sub

Public Sub DeleteSection()
    'Purpose:       to delete a section from the given INI file
    mlLastRetCode = WritePrivateProfileString(msSection, 0&, 0&, msPath)
End Sub

Public Sub EnumerateCurrentSection(ByRef sKey() As String, ByRef lCount As Long)
    'Purpose:       enumerates all of the keys in the current section of the ini file
    'Parameters:    sKey() - an array of keys that we want to enumerate
    '               lCount - the number of keys that the Section contains
    'NOTE:          since these parameters are pass ByRef this routine changes them
    '               with respect to the calling routine.  In this way we can determine
    '               if adding/deleting a key to a specific section was successful or not
    Dim sSection As String
    Dim lPos As Long
    Dim lNextPos As Long
    Dim sCur As String
    
    'set lCount to 0 and erase the array of keys
    lCount = 0
    Erase sKey
    
    'set our temprary section variable to the current INISection
    sSection = INISection
    
    If (Len(sSection) > 0) Then
        lPos = 1
        lNextPos = InStr(lPos, sSection, Chr$(0))
        Do While lNextPos <> 0
            sCur = Mid$(sSection, lPos, (lNextPos - lPos))
            If (sCur <> Chr$(0)) Then
                lCount = lCount + 1
                'new key found add this key to the array
                ReDim Preserve sKey(1 To lCount) As String
                sKey(lCount) = Mid$(sSection, lPos, (lNextPos - lPos))
                lPos = lNextPos + 1
                lNextPos = InStr(lPos, sSection, Chr$(0))
            End If
        Loop
    End If
End Sub

Public Sub EnumerateAllSections(ByRef sSections() As String, ByRef lCount As Long)
    'Purpose:       enumerates all of the sections in the current ini file
    'Parameters:    sSections() - an array of sections that we want to enumerate
    '               lCount - the number of sections that the INI file contains
    'NOTE:          since these parameters are pass ByRef this routine changes them
    '               with respect to the calling routine.  In this way we can determine
    '               if adding/deleting a section to a specific file was successful or not
    Dim sFile As String
    Dim lPos As Long
    Dim lNextPos As Long
    Dim sCur As String
    
    lCount = 0
    Erase sSections
    sFile = Sections
    If (Len(sFile) > 0) Then
        lPos = 1
        lNextPos = InStr(lPos, sFile, Chr$(0))
        Do While lNextPos <> 0
            If (lNextPos <> lPos) Then
                sCur = Mid$(sFile, lPos, (lNextPos - lPos))
                lCount = lCount + 1
                ReDim Preserve sSections(1 To lCount) As String
                sSections(lCount) = sCur
            End If
            lPos = lNextPos + 1
            lNextPos = InStr(lPos, sFile, Chr$(0))
        Loop
    End If

End Sub

Public Sub SaveFormPosition(ByRef frmThis As Object)
    'Purpose:       to save the position and size of a current form to the INI file
    'Parameters:    frmThis - a form object that we want to save information for
    Dim sSaveKey As String
    Dim sSaveDefault As String
    
    On Error GoTo SaveError
    
    sSaveKey = Key
    If Not (frmThis.WindowState = vbMinimized) Then
        Key = "Maximiszed"
        Value = (frmThis.WindowState = vbMaximized) * -1
        If (frmThis.WindowState <> vbMaximized) Then
            Key = "Left"
            Value = frmThis.Left
            Key = "Top"
            Value = frmThis.Top
            Key = "Width"
            Value = frmThis.Width
            Key = "Height"
            Value = frmThis.Height
        End If
    End If
    Key = sSaveKey
    Exit Sub
SaveError:
    Key = sSaveKey
    mlLastRetCode = 0
    Exit Sub
End Sub

Public Sub LoadFormPosition( _
        ByRef frmThis As Object, _
        Optional ByRef lMinWidth = 3000, _
        Optional ByRef lMinHeight = 3000)
    'Purpose:       to display a given form with the settings that were saved by calling
    '               the SaveFormPosition routine
    'Parameters:    frmThis - the form that we want to display values for
    '               lMinWidth - if the width parameter is not found in the saved settings
    '                   or if the information doesn't exist the form will be shown with a
    '                   default width of 3000 twips
    '               lMinHeight - if the height parameter is not found in the saved
    '                   settings or the information doesn't exist the form will be shown
    '                   with a default height of 3000 twips
    Dim sSaveKey As String
    Dim sSaveDefault As String
    Dim lLeft As Long
    Dim lTop As Long
    Dim lWidth As Long
    Dim lHeight As Long
    
    On Error GoTo LoadError
    
    sSaveKey = Key
    sSaveDefault = Default
    Default = "FAIL"
    Key = "Left"
    lLeft = CLngDefault(Value, frmThis.Left)
    
    Key = "Top"
    lTop = CLngDefault(Value, frmThis.Top)
    
    Key = "Width"
    lWidth = CLngDefault(Value, frmThis.Width)
    If (lWidth < lMinWidth) Then
        lWidth = lMinWidth
    End If
    
    Key = "Height"
    lHeight = CLngDefault(Value, frmThis.Height)
    If (lHeight < lMinHeight) Then
        lHeight = lMinHeight
    End If
    
    If (lLeft < 4 * Screen.TwipsPerPixelX) Then
        lLeft = 4 * Screen.TwipsPerPixelX
    End If
    
    If (lTop < 4 * Screen.TwipsPerPixelY) Then
        lTop = 4 * Screen.TwipsPerPixelY
    End If
    
    If (lLeft + lWidth > Screen.Width - 4 * Screen.TwipsPerPixelX) Then
        lLeft = Screen.Width - 4 * Screen.TwipsPerPixelX - lWidth
        If (lLeft < 4 * Screen.TwipsPerPixelX) Then
            lLeft = 4 * Screen.TwipsPerPixelX
        End If
        If (lLeft + lWidth > Screen.Width - 4 * Screen.TwipsPerPixelX) Then
            lWidth = Screen.Width - lLeft - 4 * Screen.TwipsPerPixelX
        End If
    End If
    
    If (lTop + lHeight > Screen.Height - 4 * Screen.TwipsPerPixelY) Then
        lTop = Screen.Height - 4 * Screen.TwipsPerPixelY - lHeight
        If (lTop < 4 * Screen.TwipsPerPixelY) Then
            lTop = 4 * Screen.TwipsPerPixelY
        End If
        If (lTop + lHeight > Screen.Height - 4 * Screen.TwipsPerPixelY) Then
            lHeight = Screen.Height - lTop - 4 * Screen.TwipsPerPixelY
        End If
    End If
    
    If (lWidth >= lMinWidth) And (lHeight >= lMinHeight) Then
        frmThis.Move lLeft, lTop, lWidth, lHeight
    End If
    
    Key = "Maximised"
    If (CLngDefault(Value, 0) <> 0) Then
        frmThis.WindowState = vbMaximized
    End If
    
    Key = sSaveKey
    Default = sSaveDefault
    Exit Sub

LoadError:
    Key = sSaveKey
    Default = sSaveDefault
    mlLastRetCode = 0
    Exit Sub
End Sub

Public Function CLngDefault( _
        ByVal sString As String, _
        Optional ByVal lDefault As Long = 0) As Long
    'Purpose:       to convert a numeric string value to a long
    'Parameters:    sString - the string to convert
    '               lDefault - the default value to return in case the passed string
    '                   is not a valid numeric string
    'Returns:       a long that equates to the conversion of the numeric string,
    '               ie, "1325" will return 1325
    Dim lRet As Long
    
    On Error Resume Next
    
    lRet = CLng(sString)
    
    If (Err.Number <> 0) Then
        CLngDefault = lDefault
    Else
        CLngDefault = lRet
    End If
End Function

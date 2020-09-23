VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   11340
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   5460
      Top             =   4410
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picInstructions 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   11340
      TabIndex        =   6
      Top             =   4725
      Width           =   11340
   End
   Begin MSComctlLib.ListView lvwFiles 
      Height          =   4215
      Left            =   105
      TabIndex        =   5
      Top             =   105
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   7435
      View            =   3
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.PictureBox picCommand 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4725
      Left            =   8070
      ScaleHeight     =   4725
      ScaleWidth      =   3270
      TabIndex        =   0
      Top             =   0
      Width           =   3270
      Begin VB.OptionButton optDBUtil 
         Caption         =   "Access 2002"
         Height          =   225
         Index           =   5
         Left            =   105
         TabIndex        =   12
         Top             =   2100
         Width           =   1485
      End
      Begin VB.OptionButton optDBUtil 
         Caption         =   "Access 2000"
         Height          =   225
         Index           =   4
         Left            =   105
         TabIndex        =   11
         Top             =   1785
         Width           =   1380
      End
      Begin VB.OptionButton optDBUtil 
         Caption         =   "Access 97"
         Height          =   225
         Index           =   3
         Left            =   105
         TabIndex        =   10
         Top             =   1470
         Width           =   1380
      End
      Begin VB.OptionButton optDBUtil 
         Caption         =   "Access 95"
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   9
         Top             =   1155
         Width           =   1380
      End
      Begin VB.OptionButton optDBUtil 
         Caption         =   "Access 2.0"
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   8
         Top             =   840
         Width           =   1275
      End
      Begin VB.OptionButton optDBUtil 
         Caption         =   "Compact"
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   7
         Top             =   210
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   435
         Left            =   1995
         TabIndex        =   4
         Top             =   105
         Width           =   1170
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   435
         Left            =   1995
         TabIndex        =   3
         Top             =   630
         Width           =   1170
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   435
         Left            =   1995
         TabIndex        =   2
         Top             =   1155
         Width           =   1170
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   435
         Left            =   1995
         TabIndex        =   1
         Top             =   1680
         Width           =   1170
      End
      Begin VB.Label lblConvert 
         Alignment       =   2  'Center
         Caption         =   "Convert To:"
         Height          =   330
         Left            =   105
         TabIndex        =   13
         Top             =   525
         Width           =   1170
      End
   End
   Begin MSComDlg.CommonDialog cdlFiles 
      Left            =   7875
      Top             =   2730
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'Module:        frmMain - Form Module
'Filename:      frmMain.frm
'Author:        Jim Kahl
'Purpose:       This serves as the main interface between the user and the target PC that
'               is used to select ActiveX Components that we currently have available for
'               testing purposes.
'NOTE:          see DevNotes.rtf for more information
'****************************************************************************************
Option Explicit
Option Compare Text

'****************************************************************************************
'API FUNCTIONS
'****************************************************************************************
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

'****************************************************************************************
'ENUMERATED CONSTANTS - PRIVATE
'****************************************************************************************
Private Enum ListviewColumns
    lvcFiles = 1
    lvcPath = 2
    #If False Then
        Private lvcFiles
        Private lvcPath
    #End If
End Enum

'****************************************************************************************
'CONSTANTS - PRIVATE
'****************************************************************************************
Private Const mBorder As Long = 100
Private Const mFileMask As String = "Access Database (*.mdb)|*.mdb|"
Private Const mStartDir As String = "C:\Program Files\Microsoft Office\Office10\Samples"

'****************************************************************************************
'VARIABLES - PRIVATE
'****************************************************************************************
Private mcIni As New cInifile

'****************************************************************************************
'EVENTS - PRIVATE
'****************************************************************************************
Private Sub cmdAdd_Click()
    'Purpose:       to allow the user a way to add another reference in the list
    Dim sFile As String
    Dim sKeys() As String
    Dim lCount As Long
    Dim lIdx As Long
    
    On Error GoTo ErrHandler
    
    With cdlFiles
        'set up the common dialog
        .DialogTitle = "Add File"
        .Filter = mFileMask
        .FilterIndex = 1
        .InitDir = mStartDir
        'raise error if user cancels dialog
        .CancelError = True
        .ShowOpen
        'retrieve the filename
        sFile = .Filename
    End With
    
    With mcIni
        .Section = "Files"
        .EnumerateCurrentSection sKeys(), lCount
        For lIdx = LBound(sKeys) To UBound(sKeys)
            .Key = sKeys(lIdx)
            'check the chosen filename to see if it already exists the ini file
            If .Value = sFile Then
                MsgBox "This item is already in the list.", vbInformation, _
                        App.ProductName
                Exit Sub
            End If
        Next lIdx
        
        'file was not found in the list so add it to ini
        If .Key = "1" And .Value = vbNullString Then
            .Key = "1"
            .Value = sFile
        Else
            .Key = CLng(sKeys(UBound(sKeys))) + 1
            .Value = sFile
        End If
    End With
    
    populateListview
    
ExitProc:
    'clean up
    Exit Sub
    
ErrHandler:
    If Err.Number = 9 Then
        'subscript out of range - occurs when there are no items in the listview control
        Resume Next
    
    'error 32755 is raised when user cancels the dialog box
    ElseIf Err.Number <> 32755 Then
        'user did not cancel so display error message
        Dim sMsg As String
        Dim sTitle As String
        Dim lRet As Long
        
        sMsg = Err.Number & ": " & Err.Description
        sTitle = "File Open Error"
        
        lRet = MsgBox(sMsg, vbCritical + vbRetryCancel, sTitle)
        If lRet = vbRetry Then
            Call cmdAdd_Click
        End If
    End If
    Resume ExitProc
End Sub

Private Sub cmdCancel_Click()
    'Purpose:       provide a way to close the form without making any changes to the
    '               user PC registry
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'Purpose:       compact and repair the selected database
    Dim cDB As New cDBUtil
    Dim lType As AcFileFormat
    
    With cDB
        .DBPath = lvwFiles.SelectedItem.Key
        Select Case True
            Case optDBUtil(0).Value
                .Compact
                GoTo ExitProc
            Case optDBUtil(1).Value
                lType = acFileFormatAccess2
            Case optDBUtil(2).Value
                lType = acFileFormatAccess95
            Case optDBUtil(3).Value
                lType = acFileFormatAccess97
            Case optDBUtil(4).Value
                lType = acFileFormatAccess2000
            Case optDBUtil(5).Value
                lType = acFileFormatAccess2002
        End Select
        .Convert lType
    End With
    
ExitProc:
    Set cDB = Nothing
End Sub

Private Sub cmdRemove_Click()
    'Purpose:       to allow the user a way to remove an item from the list
    'NOTE:          this removes it from the ini file so it won't show up the next time
    '               the application is started.
    Dim lIdx As Long
    Dim sKeys() As String       'array of key strings
    Dim lCount As Long          'count of all keys
    Dim lKey As Long            'key counter
    Dim sFile As Long
    Dim sTemp As String         'string data from the list box
    Dim sItems() As String      'items to remove from the list box
    Dim lRet As VbMsgBoxResult
    
    'if there are no items selected then we need to display the info to the user
    If lvwFiles.SelectedItem = vbNullString Then
        'if we get here then no items were selected to display a message and exit
        lRet = MsgBox("There are no items selected", vbQuestion, _
                App.ProductName)
        Exit Sub
    End If
    
    'make sure the user acually wants to remove items
    lRet = MsgBox("Do you want to remove the selected item from the list?", _
            vbQuestion + vbYesNo, App.ProductName)
    
    If lRet = vbNo Then
        Exit Sub
    End If
    
    'size the array to the number of items in the listview
    ReDim Preserve sItems(1 To lvwFiles.ListItems.Count)
    
    'cycle through and remove items that are selected
    For lIdx = 1 To lvwFiles.ListItems.Count
        If lvwFiles.ListItems(lIdx).Selected Then
            'split the string data from the list box
            'for purposes here we only care about the filename anyway
            sTemp = lvwFiles.ListItems(lIdx).Text
            With mcIni
                .Section = "Files"
                .EnumerateCurrentSection sKeys(), lCount
                For lKey = LBound(sKeys) To UBound(sKeys)
                    .Key = sKeys(lKey)
                    'check the filename for the item selected and match it to the
                    'filename in the ini file
                    If sTemp = getFilename(.Value) Then
                        'delete the key from the ini file
                        .DeleteKey
                        'set the array element that matches the item in the list to be
                        'value of the items index
                        sItems(lIdx) = lIdx
                    End If
                Next lKey
            End With
        End If
    Next lIdx
    
    populateListview
    
End Sub

Private Sub Form_Load()
    
    'initialize the form
    Me.Caption = App.ProductName & Space$(1) & _
            "v" & App.Major & "." & App.Minor & "." & App.Revision
    picInstructions.AutoRedraw = True
    picInstructions.Print "You can only choose one database at a time."
    populateListview
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim lIdx As Long
    Dim sKeys() As String
    Dim sItems() As String
    Dim lCount As Long
    
    On Error Resume Next
    
    'before we close down the app let's clean up the ini file
    With mcIni
        .Section = "Files"
        .EnumerateCurrentSection sKeys(), lCount
        ReDim sItems(1 To lCount)
        For lIdx = 1 To lCount
            'store the filespecs in our temporary array and then delete it from the
            'ini file
            .Key = sKeys(lIdx)
            sItems(lIdx) = .Value
            .DeleteKey
        Next lIdx
        
        'now recreate the ini file so the numbers are sequential
        'NOTE: this is important when trying to add new files to the list because we
        'don't need to check for an open number we just take the next number after the
        'last one in the ini file.
        'if there is no count then make sure there is an item in the list
        If lCount = 0 Then
            .Section = "Files"
            .Key = "1"
            .Value = ""
        End If
        
        For lIdx = 1 To lCount
            .Key = lIdx
            .Value = sItems(lIdx)
        Next lIdx
    End With
    
    'perform cleanup
    Set mcIni = Nothing
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    
    'resize the listbox when the form gets resized
    With lvwFiles
        .Width = picCommand.Left - (mBorder * 2)
        .Height = Me.Height - .Top - picInstructions.Height - mBorder * 5
    End With
End Sub

Private Sub lvwFiles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ColumnHeader.Key = "Files" Then
        If lvwFiles.SortOrder = lvwAscending Then
            lvwFiles.SortOrder = lvwDescending
        Else
            lvwFiles.SortOrder = lvwAscending
        End If
    End If
End Sub

'****************************************************************************************
'METHODS - PRIVATE
'****************************************************************************************
Private Function getFilename(ByRef FileSpec As String) As String
    'Returns:       filename and extension of the file
    Dim lPos As Long
    
    lPos = InStrRev(FileSpec, "\")
    
    getFilename = Right$(FileSpec, Len(FileSpec) - lPos)
End Function

Private Sub populateListview()
    'Purpose:       reads information from the pgaRegister.ini file and then reads each
    '               of the files version info into the FixedFileInfo class definition to
    '               populate the listview control with information obtained from the
    '               class properties
    Dim cFI As New cFixedFileInfo
    Dim sKeys() As String
    Dim sName As String
    Dim lCount As Long
    Dim lIdx As Long
    Dim hWnd As Long
    Dim lRet As Long
    Dim li As ListItem
    Static bDone As Boolean
    
    Screen.MousePointer = vbHourglass
    
    'set up and read the ini file items we need
    With mcIni
        .Path = App.Path & "\pgadbutil.ini"
        .Section = "Files"
        .EnumerateCurrentSection sKeys(), lCount
    End With
    
    LockWindowUpdate Me.hWnd
    
    'initialize the listview control
    With lvwFiles
        'set up the properties as we want them
        .ListItems.Clear
        .AllowColumnReorder = False
        .MultiSelect = False
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .LabelEdit = lvwManual
        .Sorted = True
        .SortOrder = lvwAscending
        If Not bDone Then
            'set up the columns
            .ColumnHeaders.Add
            .ColumnHeaders.Add
            With .ColumnHeaders(lvcFiles)
                .Text = "File"
                .Width = 2000
                .Key = .Text
            End With
            With .ColumnHeaders(lvcPath)
                .Text = "Path"
                .Width = lvwFiles.Width - _
                        lvwFiles.ColumnHeaders(lvcFiles).Width - _
                        Screen.TwipsPerPixelX
                .Key = .Text
            End With
            bDone = True
        End If
    End With
    
    'loop through and add items to the list box
    For lIdx = LBound(sKeys) To UBound(sKeys)
        Set cFI = New cFixedFileInfo
        With cFI
            mcIni.Key = sKeys(lIdx)
            sName = mcIni.Value
            .FullPathName = sName
            Set li = lvwFiles.ListItems.Add()
            li.SmallIcon = 1
            li.Key = sName
            li.Text = .Filename
            li.SubItems(1) = sName
        End With
    Next lIdx
    
    LockWindowUpdate 0&
    
    Screen.MousePointer = vbDefault
    
    Set cFI = Nothing
End Sub

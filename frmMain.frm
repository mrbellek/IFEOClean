VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "IFEO Clean"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleMode       =   0  'User
   ScaleWidth      =   8000
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHelp 
      BackColor       =   &H8000000F&
      Height          =   615
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   3240
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "What does this mean?"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Tag             =   "What does this mean?"
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdFix 
      Caption         =   "Delete selected debuggers"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4048
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupCopy 
         Caption         =   "Copy to clipboard"
      End
      Begin VB.Menu mnuPopupSave 
         Caption         =   "Save to disk..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function SHDeleteKey Lib "shlwapi.dll" Alias "SHDeleteKeyA" (ByVal lRootKey As Long, ByVal szKeyToDelete As String) As Long

Private Declare Function SHFileExists Lib "shell32" Alias "#45" (ByVal szPath As String) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800

Private Const VER_PLATFORM_WIN32_NT = 2

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003

Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_SZ = 1

Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

Private sWinDir$, sSysDir$, bIsWinNT As Boolean

Private Sub cmdFix_Click()
    Dim i&, sName$
    Const sIFEOkey$ = "Software\Microsoft\Windows NT\CurrentVersion\Image File Execution Options"
    For i = 1 To lvwMain.ListItems.Count
        If lvwMain.ListItems(i).Checked Then
            sName = lvwMain.ListItems(i).Text
            RegDelValue HKEY_LOCAL_MACHINE, sIFEOkey & "\" & sName, "Debugger"
            If RegValExists(HKEY_LOCAL_MACHINE, sIFEOkey & "\" & sName, "Debugger") Then
                lvwMain.ListItems(i).SubItems(2) = "Fix FAILED!"
            Else
                If Not RegKeyHasValues(HKEY_LOCAL_MACHINE, sIFEOkey & "\" & sName) Then
                    RegDelKey HKEY_LOCAL_MACHINE, sIFEOkey & "\" & sName
                End If
                lvwMain.ListItems(i).SubItems(2) = "Fixed"
            End If
        End If
    Next i
End Sub

Private Sub cmdHelp_Click()
    If cmdHelp.Caption <> "Hide help text" Then
        txtHelp.Visible = True
        cmdHelp.Caption = "Hide help text"
        Form_Resize
    Else
        txtHelp.Visible = False
        cmdHelp.Caption = cmdHelp.Tag
        Form_Resize
    End If
End Sub

Private Sub cmdScan_Click()
    Dim hKey&, i&, sName$, sDebugger$
    Const sCriticalFiles$ = "userinit.exe|regedit.exe|explorer.exe|regedt32.exe|taskmgr.exe|winlogon.exe"
    Const sIFEOkey$ = "Software\Microsoft\Windows NT\CurrentVersion\Image File Execution Options"
    lvwMain.ListItems.Clear
    
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, sIFEOkey, 0, KEY_ENUMERATE_SUB_KEYS, hKey) <> 0 Then
        'could not open key
        Exit Sub
    End If
    
    sName = String(255, 0)
    If RegEnumKeyEx(hKey, 0, sName, Len(sName), 0, vbNullString, 0, ByVal 0) <> 0 Then
        'no subkeys
        RegCloseKey hKey
        Exit Sub
    End If
    
    Do
        sName = TrimNull(sName)
        
        sDebugger = RegGetString(HKEY_LOCAL_MACHINE, sIFEOkey & "\" & sName, "Debugger")
        If sDebugger <> vbNullString And sName <> "Your Image File Name Here without a path" Then
            lvwMain.ListItems.Add , "key" & i, sName
            lvwMain.ListItems("key" & i).SubItems(1) = sDebugger
            If InStr(1, sCriticalFiles, sName, vbTextCompare) > 0 Then
                lvwMain.ListItems("key" & i).Bold = True
            End If
            
            If Mid(sDebugger, 1, 1) = """" Then
                'surrounded by quotes, easy
                sDebugger = Mid(sDebugger, 2)
                sDebugger = Left(sDebugger, InStr(sDebugger, """") - 1)
            Else
                'if it has spaces, it might have parameters
                If InStr(sDebugger, " ") > 0 Then
                    If InStr(1, sDebugger, ".exe", vbTextCompare) > 0 Then
                        'cut off everything after .exe
                        sDebugger = Left(sDebugger, InStr(1, sDebugger, ".exe", vbTextCompare) + 3)
                    Else
                        'cut off everything after last space (risky)
                        sDebugger = Left(sDebugger, InStrRev(sDebugger, " ") - 1) & ".exe"
                    End If
                Else
                    If InStr(1, sDebugger, ".exe", vbTextCompare) = 0 Then
                        sDebugger = sDebugger & ".exe"
                    End If
                End If
                'if it's not a full path, guess it
                If InStr(sDebugger, "\") = 0 Then
                    If FileExists(App.Path & "\" & sDebugger) Then
                        sDebugger = App.Path & "\" & sDebugger
                    Else
                        If FileExists(sWinDir & "\" & sDebugger) Then
                            sDebugger = sWinDir & "\" & sDebugger
                        Else
                            If FileExists(sSysDir & "\" & sDebugger) Then
                                sDebugger = sSysDir & "\" & sDebugger
                            End If
                        End If
                    End If
                End If
            End If
            If FileExists(sDebugger) Then
                lvwMain.ListItems("key" & i).SubItems(2) = "File exists"
            Else
                lvwMain.ListItems("key" & i).SubItems(2) = "File not found!"
                lvwMain.ListItems("key" & i).ForeColor = vbRed
                lvwMain.ListItems("key" & i).Checked = True
            End If
        End If
        
        sName = String(255, 0)
        i = i + 1
    Loop Until RegEnumKeyEx(hKey, i, sName, Len(sName), 0, vbNullString, 0, ByVal 0) <> 0
    RegCloseKey hKey
End Sub

Private Sub Form_Load()
    With lvwMain
        .View = lvwReport
        .FullRowSelect = True
        .Checkboxes = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , "app", "Host program file", 3000
        .ColumnHeaders.Add , "hook", "Debugger program file", 3000
        .ColumnHeaders.Add , "status", "Status", 2500
    End With
    
    Me.Width = 8000
    Me.Height = 6000
    
    CheckWindows
        
    If Not bIsWinNT Then
        MsgBox "This program addresses an issue that only converns Windows NT4, " & _
               "Windows 2000, Windows XP, Windows Vista and newer. Your Windows " & _
               "version is not affected by this issue.", vbInformation
    End If
    
    txtHelp.Text = "Each Windows program can be setup to have a different program " & _
                   "start together with it, from the 'Image File Execution Options' " & _
                   "(IFEO) Registry key. This program acts as a debugger for the host " & _
                   "program. If this debugger program is deleted, the host program " & _
                   "will not start! This is very serious if a debugger is attached " & _
                   "to system processes but has been deleted, possibly rendering " & _
                   "a system useless. Only after the debugger has been disabled " & _
                   "will the host program start normally." & vbCrLf & vbCrLf & _
                   "All registered debuggers are listed here: those in BOLD " & _
                   "are system processes, those in RED are missing (which means " & _
                   "the host program will not start). It is recommended to fix " & _
                   "any items in red, and it is critical to fix items in bold red!"
End Sub

Private Function TrimNull$(s$)
    If InStr(s, Chr(0)) > 0 Then
        TrimNull = Left(s, InStr(s, Chr(0)) - 1)
    Else
        TrimNull = s
    End If
End Function

Private Function RegGetString$(lHive&, sKey$, sValue$)
    Dim hKey&, uData() As Byte, sData$, lDataLen&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_QUERY_VALUE, hKey) = 0 Then
        ReDim uData(0)
        RegQueryValueEx hKey, sValue, 0, ByVal 0, uData(0), lDataLen
        ReDim uData(lDataLen)
        If RegQueryValueEx(hKey, sValue, 0, ByVal 0, uData(0), lDataLen) = 0 Then
            sData = TrimNull(StrConv(uData, vbUnicode))
            RegGetString = sData
        End If
        RegCloseKey hKey
    End If
End Function

Private Function RegKeyExists(lHive&, sKey$) As Boolean
    Dim hKey&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_QUERY_VALUE, hKey) = 0 Then
        RegKeyExists = True
        RegCloseKey hKey
    End If
End Function

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    On Error Resume Next
    
    'cmdHelp.Left = Me.ScaleWidth - 2175
    lvwMain.Width = Me.ScaleWidth - 240 - 60
    txtHelp.Width = Me.ScaleWidth - 240 - 60
    If txtHelp.Visible Then
        lvwMain.Height = Me.ScaleHeight - 1815 - 240
        txtHelp.Top = lvwMain.Top + lvwMain.Height + 120
        txtHelp.Height = Me.ScaleHeight - 960 - 120 - lvwMain.Height
    Else
        lvwMain.Height = Me.ScaleHeight - 960
    End If
End Sub

Private Function FileExists(sFile$) As Boolean
    FileExists = False
    If bIsWinNT Then
        If SHFileExists(StrConv(sFile, vbUnicode)) Then FileExists = True
    Else
        If SHFileExists(sFile) Then FileExists = True
    End If
End Function

Private Sub CheckWindows()
    sWinDir = String(260, 0)
    sWinDir = Left(sWinDir, GetWindowsDirectory(sWinDir, Len(sWinDir)))
    sSysDir = String(260, 0)
    sSysDir = Left(sSysDir, GetSystemDirectory(sSysDir, Len(sSysDir)))
        
    Dim uOVI As OSVERSIONINFO
    uOVI.dwOSVersionInfoSize = Len(uOVI)
    GetVersionEx uOVI
    If uOVI.dwPlatformId = VER_PLATFORM_WIN32_NT Then bIsWinNT = True
End Sub

Private Sub RegDelValue(lHive&, sKey$, sValue$)
    Dim hKey&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_WRITE, hKey) = 0 Then
        RegDeleteValue hKey, sValue
        RegCloseKey hKey
    End If
End Sub

Private Function RegValExists(lHive&, sKey$, sValue$) As Boolean
    Dim hKey&, uData() As Byte
    If RegOpenKeyEx(lHive, sKey, 0, KEY_QUERY_VALUE, hKey) = 0 Then
        ReDim uData(260)
        If RegQueryValueEx(hKey, sValue, 0, ByVal 0, uData(0), UBound(uData)) = 0 Then
            RegValExists = True
        End If
        RegCloseKey hKey
    End If
End Function

Private Function RegKeyHasValues(lHive&, sKey$) As Boolean
    Dim hKey&, sName$, uData() As Byte
    If RegOpenKeyEx(lHive, sKey, 0, KEY_QUERY_VALUE, hKey) = 0 Then
        sName = String(65536, 0)
        ReDim uData(260)
        If RegEnumValue(hKey, 0, sName, Len(sName), 0, ByVal 0, uData(0), UBound(uData)) = 0 Then
            RegKeyHasValues = True
        End If
        RegCloseKey hKey
    End If
End Function

Private Sub RegDelKey(lHive&, sKey$)
    SHDeleteKey lHive, sKey
End Sub

Private Sub lvwMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And lvwMain.ListItems.Count > 0 Then PopupMenu mnuPopup
End Sub

Private Sub mnuPopupCopy_Click()
    Dim i&, sMsg$
    For i = 1 To lvwMain.ListItems.Count
        sMsg = sMsg & vbCrLf & lvwMain.ListItems(i).Text & " - " & lvwMain.ListItems(i).SubItems(1) & " [" & lvwMain.ListItems(i).SubItems(2) & "]"
    Next i
    If sMsg = vbNullString Then Exit Sub
    sMsg = Mid(sMsg, 2)
    Clipboard.Clear
    Clipboard.SetText sMsg
    MsgBox "Copied scan results to your clipboard.", vbInformation
End Sub

Private Sub mnuPopupSave_Click()
    Dim uOFN As OPENFILENAME, sFile$, i&, sLog$
    With uOFN
        .lStructSize = Len(uOFN)
        .flags = OFN_HIDEREADONLY Or OFN_NONETWORKBUTTON Or OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
        .lpstrDefExt = "txt"
        .lpstrFile = String(260, 0)
        .lpstrFilter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        .lpstrFilter = Replace(.lpstrFilter, "|", Chr(0)) & Chr(0) & Chr(0)
        .lpstrInitialDir = App.Path
        .lpstrTitle = "Save scan results..."
        .nMaxFile = Len(.lpstrFile)
        GetSaveFileName uOFN
        sFile = TrimNull(.lpstrFile)
    End With
    If sFile = vbNullString Then Exit Sub
    
    For i = 1 To lvwMain.ListItems.Count
        sLog = sLog & vbCrLf & lvwMain.ListItems(i).Text & " - " & lvwMain.ListItems(i).SubItems(1) & " [" & lvwMain.ListItems(i).SubItems(2) & "]"
    Next i
    If sLog = vbNullString Then Exit Sub
    sLog = Mid(sLog, 2)
    sLog = "IFEOClean v" & App.Major & "." & Format(App.Minor, "00") & "." & App.Revision & _
           vbCrLf & "Logfile saved on " & Date & " at " & Time & vbCrLf & vbCrLf & _
           "Listing all debuggers on local system:" & vbCrLf & sLog
    
    Open sFile For Output As #1
        Print #1, sLog
    Close #1
    MsgBox "Saved scan results to disk:" & vbCrLf & sFile, vbInformation
End Sub

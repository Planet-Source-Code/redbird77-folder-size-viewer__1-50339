Attribute VB_Name = "mBrowse"
' file    : mBrowse.bas
' revised : 2001-04-17
' author  : redbird77
' email   : redbird77@earthlink.net
' www     : http://home.earthlink.net/~redbird77

Option Explicit

Private m_sPreSelDir As String

Private Const MAX_PATH As Long = 260
Private Const WM_USER  As Long = &H400

Public Enum BrowseInfoFlags
    BIF_BROWSEFORCOMPUTER = &H1000
    BIF_BROWSEFORPRINTER = &H2000
    BIF_BROWSEINCLUDEFILES = &H4000
    BIF_DONTGOBELOWDOMAIN = &H2
    BIF_EDITBOX = &H10
    BIF_RETURNFSANCESTORS = &H8
    BIF_RETURNONLYFSDIRS = &H1
    BIF_STATUSTEXT = &H4
    BIF_VALIDATE = &H20
End Enum

' See the function BrowseCallbackProc for more comments on these messages.
Public Enum BrosweForFolderMessages

    ' Messages that define events.
    BFFM_SELCHANGED = &H2
    BFFM_INITIALIZED = &H1

    ' Messages that the callback function can send to the dialog.
    BFFM_SETSTATUSTEXTA = (WM_USER + 100)
    BFFM_ENABLEOK = (WM_USER + 101)
    BFFM_SETSELECTIONA = (WM_USER + 102)
    
End Enum

Private Type BrowseInfo
    hwndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As String
    lpszTitle      As String
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectoryA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Function BrowseForFolder(ByVal hwnd As Long, _
                                Optional ByVal sTitle As String = "Select a folder.", _
                                Optional ByVal lFlags As BrowseInfoFlags = BIF_RETURNONLYFSDIRS, _
                                Optional ByVal sPreSelDir As String = "") As String

    Dim BI As BrowseInfo, sDir As String
    
    If sPreSelDir <> "" Then
        m_sPreSelDir = sPreSelDir
    'Else
    '    m_sPreSelDir = CurDir$()
    End If
    
    With BI
    
        ' Set owner of Browse dialog box.  If this is zero,
        ' then the dialog is display non-modally.
        .hwndOwner = hwnd
        
        .lpszTitle = sTitle
        
        .lpfnCallback = GetAddress(AddressOf BrowseCallbackProc)
        
        .ulFlags = lFlags
        
        .pIDLRoot = 0&
        
    End With
    
    ' <From VB6 Help File Re: SHBrowseForFolder>
    '
    ' Returns the address of an item identifier list that specifies
    ' the location of the selected folder relative to the root of
    ' the namespace. If the user chooses the Cancel button in the
    ' dialog box, the return value is NULL.
    
    sDir = GetFolderPathFromID(SHBrowseForFolder(BI))
    
    If sDir <> "" Then sDir = sDir & IIf(Right$(sDir, 1) = "\", "", "\")
    
    BrowseForFolder = sDir

    ' Return value is the user selected folder, "" if user canceled.
    
End Function

Private Function GetFolderPathFromID(ByVal lpIDL As Long) As String

    Dim sPath As String
    Dim iPos  As Integer
    
    ' If user cancelled then GetFolderPathFromID = "".
    If lpIDL = 0 Then Exit Function
    
    ' Fill buffer with nulls.
    sPath = String$(MAX_PATH, vbNullChar)
    
    ' Get folder path.
    SHGetPathFromIDList lpIDL, sPath
    
    CoTaskMemFree lpIDL
    
    ' Return the part before the null terminator.
    iPos = InStr(sPath, vbNullChar)
    If iPos Then sPath = Left$(sPath, iPos - 1)
    
    GetFolderPathFromID = sPath
        
End Function

Private Function BrowseCallbackProc(ByVal hwnd As Long, _
                                    ByVal lMsg As Long, _
                                    ByVal lParam As Long, _
                                    ByVal lpData As Long) As Long
    Dim sBuf As String
    Dim lLen As Long
    
    Select Case lMsg
    
        ' -------------------------------------------------------------------
        ' BFFM_INITIALIZED
        ' -------------------------------------------------------------------
        ' Indicates the browse dialog box has finished initializing. The
        ' lParam parameter is NULL. (msdn)
        
        Case BFFM_INITIALIZED

            'Debug.Print "BFFM_INITIALIZED: "; Hex$(lMsg)

            ' ---------------------------------------------------------------
            ' BFFM_SETSELECTIONA
            ' ---------------------------------------------------------------
            ' Selects the specified folder. The message's lParam is the PIDL
            ' of the folder to select if wParam is FALSE, or the path of the
            ' folder otherwise. (msdn)
            If m_sPreSelDir <> "" Then
            
                SendMessage hwnd, BFFM_SETSELECTIONA, ByVal 1&, _
                            ByVal m_sPreSelDir
            End If

        ' -------------------------------------------------------------------
        ' BFFM_SELCHANGED
        ' -------------------------------------------------------------------
        ' Indicates the selection has changed. The lParam parameter contains
        ' the address of the item identifier list for the newly selected
        ' folder. (msdn)
        
        Case BFFM_SELCHANGED

            'Debug.Print "BFFM_SELCHANGED: "; Hex$(lMsg)
            'Debug.Print "lParam: "; lParam

            ' ---------------------------------------------------------------
            ' BFFM_SETSTATUSTEXTA
            ' ---------------------------------------------------------------
            ' Sets the status text to the null-terminated string specified by
            ' the message's lParam parameter. (msdn)
        
'            SendMessage hwnd, BFFM_SETSTATUSTEXTA, ByVal 0&, _
'                        ByVal GetFolderPathFromID(lParam)

    End Select

End Function

' ---------------------------------------------------------------------------
' Helper Functions
' ---------------------------------------------------------------------------

Private Function GetAddress(ByVal lProcAddress As Long) As Long

    ' Wrapper of the AddressOf keyword to prevent syntax errors.
    GetAddress = lProcAddress
    
End Function

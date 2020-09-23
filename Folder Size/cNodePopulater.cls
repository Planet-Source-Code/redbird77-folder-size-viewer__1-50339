VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNodePopulater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const DOT     As String = "."
Private Const EXT_ALL As String = "*.*"

Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Private Const INVALID_HANDLE_VALUE     As Long = (-1)
Private Const MAX_PATH                 As Long = 260

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Type FILETIME
    dwLowDateTime  As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime   As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime  As FILETIME
    nFileSizeHigh    As Long
    nFileSizeLow     As Long
    dwReserved0      As Long
    dwReserved1      As Long
    cFileName        As String * MAX_PATH
    cAlternate       As String * 14
End Type

Private Enum SizeMethodConstants
    smInitial
    smOnClick
End Enum

Private eMethod As SizeMethodConstants

Private m_oTreeView As TreeView
Private colExpanded As Collection
Private colInfo     As Collection

Public Property Get oTreeView() As TreeView
    Set oTreeView = m_oTreeView
End Property

Public Property Set oTreeView(n As TreeView)
    Set m_oTreeView = n
End Property

Private Sub Class_Initialize()
    Set colExpanded = New Collection
    Set colInfo = New Collection
End Sub

Private Sub Class_Terminate()
    Set colInfo = Nothing
    Set colExpanded = Nothing
    Set m_oTreeView = Nothing
End Sub

Public Sub ExpandNode(n As Node)

    If n.Children = 0 Then Exit Sub

    If n.Child.Text = "FILLER" Then m_oTreeView.Nodes.Remove n.Child.Index
    
    PopulateNode n.Key
    
End Sub

Public Sub PopulateNode(ByVal sStartFolder As String)

Dim i        As Long
Dim lCount   As Long
Dim sNames() As String
Dim sPath    As String
Dim n        As Node
Dim dSize    As Double

    ' If the node has already been filled then exit sub.
    If colExpanded.Item(sStartFolder) = True Then Exit Sub

    ' Get 1-level deep subfolders of start folder.
    lCount = pGetFolders(sStartFolder, sNames())
    
    For i = 0 To lCount - 1
        ' Construct the complete path.
        sPath = sStartFolder & sNames(i) & "\"
    
        ' Retrieve size and subfolder existence from Info collection
        ' that was filled when root folder was chosen.
        dSize = Mid$(colInfo(sPath), 2)
        
        ' Add the node and set it's tag to the size.
        Set n = m_oTreeView.Nodes.Add(sStartFolder, tvwChild, sPath, sNames(i) & _
                                     " [" & pFormatSize(dSize) & "]")
        n.Tag = CStr(dSize)
        
        ' Add a FALSE entry in the Expanded collection.
        Call colExpanded.Add(False, sPath)
        
        ' Add filler node to indicate subfolders existence.
        If CBool(Left$(colInfo(sPath), 1)) Then
            Set n = m_oTreeView.Nodes.Add(sPath, tvwChild, , "FILLER")
        End If
    Next
        
    ' Mark node as expanded.  Am I missing something REALLY OBVIOUS or is
    ' there a way to change a member of a collection without removing and
    ' re-adding it when the collection is not a collection of class instances,
    ' but of simple data type?
    colExpanded.Remove sStartFolder
    colExpanded.Add True, sStartFolder
    
End Sub

Public Sub GraphNode(n As Node, Canvas As PictureBox, eOrder As SortOrderConstants)

Dim oChart     As cPieChart
Dim nChild     As Node
Dim dTotal     As Double
Dim dLoose     As Double
Dim dSizeChild As Double
Dim bDraw      As Boolean
Dim LI         As ListItem

    Set oChart = New cPieChart
    
    Set oChart.Canvas = Canvas
    
    oChart.Canvas.Font.Size = 6
    oChart.Canvas.Font.Name = "Terminal"
    
    ' Add sections.
    If n.Children Then
        Set nChild = n.Child
        
        Do
            If nChild.Text <> "FILLER" Then
                dSizeChild = CDbl(nChild.Tag)
                dTotal = dTotal + dSizeChild
                
                If dSizeChild Then
                    oChart.Sections.Add CDbl(nChild.Tag), -1, _
                                        pGetNameFromPath(nChild.Key), nChild.Key
                    bDraw = True
                End If
            End If
            
            Set nChild = nChild.Next
            
        Loop Until nChild Is Nothing
        
    End If
    
    ' Find size of loose files.  Loose files are files located in
    ' the selected folder, but not in a subfolder.
    dLoose = CDbl(n.Tag) - dTotal
        
    If dLoose Then
        oChart.Sections.Add dLoose, -1, "Loose", n.Key
        bDraw = True
    End If

Dim psize As Double, i As Integer

    If bDraw Then
        oChart.Sections.Sort soSize
        
        ' Add to listview? (maybe an Event OnSectionGraph(Section as Section))
        ' TODO: Make this optional, and ListView a property (no fMain ref).
        fMain.lvwFolders.ListItems.Clear
        For i = 1 To oChart.Sections.Count
            Set LI = fMain.lvwFolders.ListItems.Add(, oChart.Sections(i).Key, _
                                                    oChart.Sections(i).Text)
            LI.SubItems(1) = pFormatSize(oChart.Sections(i).Size)
            ' Get % of parent size of selected node.
            psize = CDbl(oTreeView.Nodes(oTreeView.SelectedItem.Key).Tag)
            LI.SubItems(2) = Format$(oChart.Sections(i).Size / psize, "0.00%")
            
            ' Get % of root size of selected node.
            LI.SubItems(3) = Format$(oChart.Sections(i).Size / oTreeView.Nodes(1).Tag, "0.00%")
        Next
        
        oChart.Draw
    Else
        ' TODO: Don't like this.  Change error-handling to within cPieChart.
        With oChart.Canvas
            .Cls
            .CurrentX = 2: .CurrentY = 2
            oChart.Canvas.Print n.Key
            oChart.Canvas.Print "contains no files or folders."
        End With
    End If
    
    Set oChart = Nothing
    
End Sub

Private Function GetFolderSize(ByVal sPath As String, _
                               ByRef bHasSubs As Boolean) As Double
                               
' Note: This function is recursive.

    Dim h      As Long
    Dim FD     As WIN32_FIND_DATA
    Dim r      As Long
    Dim dSize  As Double
    Dim sName  As String
    
    ' Get handle to first file or subfolder in folder.
    h = FindFirstFile(sPath & EXT_ALL, FD)
    bHasSubs = False

    If h <> INVALID_HANDLE_VALUE Then
        Do
            sName = Left$(FD.cFileName, InStr(FD.cFileName, vbNullChar) - 1)
            If Left$(sName, 1) <> DOT Then
                If (FD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
                    bHasSubs = True
        
                    ' If the handle is to a folder then call the function recursively.
                    dSize = dSize + GetFolderSize(sPath & sName & "\", True)
                Else
                    dSize = dSize + FD.nFileSizeLow
                End If
                
            End If
        
        Loop While FindNextFile(h, FD)
        
        r = FindClose(h): Debug.Assert r
    End If
    
    ' Return the folder size and add the size to the Collection with
    ' the folder path as the key for later referencing.
    GetFolderSize = dSize
    colInfo.Add CStr(bHasSubs * -1) & "" & dSize, sPath

End Function

Public Sub SetRoot(ByVal sRootFolder As String)

Dim n As Node, dSize As Double

    m_oTreeView.Nodes.Clear
    
    dSize = GetFolderSize(sRootFolder, 0)
    
    ' Add the root node.
    Set n = m_oTreeView.Nodes.Add(, , sRootFolder, pGetNameFromPath(sRootFolder) & _
                                 " [" & pFormatSize(dSize) & "]")
                              
    n.Tag = CStr(dSize)
    
    ' Add a FALSE entry in the Expanded collection.
    colExpanded.Add False, sRootFolder
      
End Sub

Private Function pGetFolders(ByVal sStartFolder As String, _
                            ByRef sNames() As String) As Long
                            
' Get 1-level deep subfolders of start folder.  This function is not recursive.

Dim h As Long, r As Long, FD As WIN32_FIND_DATA, n As Long

    h = FindFirstFile(sStartFolder & EXT_ALL, FD)

    If h <> INVALID_HANDLE_VALUE Then
        Do
            If Left$(FD.cFileName, 1) <> DOT Then
                If (FD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
                    ReDim Preserve sNames(n)
                    sNames(n) = Left$(FD.cFileName, InStr(FD.cFileName, vbNullChar) - 1) 'sName
                    n = n + 1
                End If
            End If

        Loop While FindNextFile(h, FD)

        r = FindClose(h): Debug.Assert r
    End If

    ' Sort subfolders (if any).
    If n Then QuickSort sNames(), 0, UBound(sNames())
    
    ' GetFolders returns number of subfolders, zero of none.
    pGetFolders = n
    
End Function

Private Function pFormatSize(ByVal dSize As Double) As String

' 1024  b = 1 kb: 1024 kb = 1 mb

    If dSize < 1024 Then
        pFormatSize = dSize & " bytes"
    Else
        dSize = dSize / 1024
        If dSize < 1000 Then
            pFormatSize = Format$(dSize, "#,##0.0") & " kb"
        Else
            pFormatSize = Format$(dSize / 1024, "#,##0.0") & " mb"
        End If
    End If
    
End Function

Private Function pGetNameFromPath(ByVal sPath As String) As String

' Assumes sPath will always end in a backslash.

Dim i As Integer, s As String

    ' Special case = when drive root (i.e. C:\, D:\, etc.)
    pGetNameFromPath = sPath
    
    For i = Len(sPath) - 1 To 1 Step -1
        If Mid$(sPath, i, 1) = "\" Then
            s = Mid$(sPath, i + 1)
            pGetNameFromPath = Left$(s, Len(s) - 1)
            Exit Function
        End If
    Next
    
End Function

Private Sub QuickSort(sVals() As String, lo As Long, hi As Long)

' Thank you Vbnet.

Dim tlo As Long, thi As Long, x As String, y As String

    tlo = lo: thi = hi
    x = sVals((lo + hi) / 2)

    While (tlo <= thi)
        While (StrComp(sVals(tlo), x, vbTextCompare) = -1 And tlo < hi)
            tlo = tlo + 1
        Wend

        While (StrComp(x, sVals(thi), vbTextCompare) = -1 And thi > lo)
           thi = thi - 1
        Wend

        If (tlo <= thi) Then
            y = sVals(tlo)
            sVals(tlo) = sVals(thi)
            sVals(thi) = y
            tlo = tlo + 1
            thi = thi - 1
        End If
    Wend

   If (lo < thi) Then QuickSort sVals, lo, thi
   If (tlo < hi) Then QuickSort sVals, tlo, hi

End Sub


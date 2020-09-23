VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Folder Size"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8865
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   8865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtRootFolder 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
   Begin MSComctlLib.TreeView trvFolders 
      Height          =   7455
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   13150
      _Version        =   393217
      Indentation     =   706
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picGraph 
      AutoRedraw      =   -1  'True
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   4935
      Left            =   3840
      ScaleHeight     =   325
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   325
      TabIndex        =   4
      Top             =   600
      Width           =   4935
   End
   Begin MSComctlLib.ListView lvwFolders 
      Height          =   2415
      Left            =   3840
      TabIndex        =   6
      Top             =   5640
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "% of Parent"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "% of Root"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.ListBox List2 
      Height          =   2400
      Left            =   3840
      TabIndex        =   5
      Top             =   5640
      Width           =   4935
   End
   Begin VB.Label lblCap 
      Caption         =   "Root Folder:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project : pFolderSize
' revised : 2003-12-05
' author  : redbird77
' email   : redbird77@earthlink.net
' www     : http://home.earthlink.net/~redbird77
' about   : See Related Document - README.txt.

Option Explicit

Private NP As cNodePopulater

Private Sub cmdBrowse_Click()

Dim t     As Single
Dim sRoot As String

    ' Get root folder.  BrowseForFolder makes sure path ends in a backslash.
    sRoot = mBrowse.BrowseForFolder(Me.hwnd, , , App.Path)
    
    ' Exit sub if user cancelled.
    If Len(sRoot) = 0 Then Exit Sub

    't = Timer

    txtRootFolder.Text = sRoot
    
    Set NP = New cNodePopulater
    
    With NP
        Set .oTreeView = trvFolders
        ' Clear all nodes and add the root node.
        .SetRoot sRoot
        ' Fill treeview with 1-level deep subfolders.
        .PopulateNode sRoot
    End With
    
    'lblTime.Caption = Format$(Timer - t, "0.00")

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set NP = Nothing
    Set fMain = Nothing
End Sub

Private Sub trvFolders_Expand(ByVal Node As MSComctlLib.Node)
    'Dim t As Single
    't = Timer
    NP.ExpandNode Node
    'lblTime.Caption = Format$(Timer - t, "0.000")
End Sub

Private Sub trvFolders_NodeClick(ByVal Node As MSComctlLib.Node)
    trvFolders_Expand Node
    NP.GraphNode Node, picGraph, soText
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private m_lColor As Long
Private m_dSize  As Double
Private m_sText  As String
Private m_sKey   As String

' -------------------------------------------------------------------

Public Property Let Size(ByVal n As Double)
    m_dSize = n
End Property

Public Property Get Size() As Double
    Size = m_dSize
End Property

' -------------------------------------------------------------------

Public Property Let Color(ByVal n As Long)
    m_lColor = n
End Property

Public Property Get Color() As Long
    Color = m_lColor
End Property

' -------------------------------------------------------------------

Public Property Let Text(ByVal n As String)
    m_sText = n
End Property

Public Property Get Text() As String
    Text = m_sText
End Property

' -------------------------------------------------------------------

Public Property Let Key(ByVal n As String)
    m_sKey = n
End Property

Public Property Get Key() As String
    Key = m_sKey
End Property


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"cSection"
Attribute VB_Ext_KEY = "Member0" ,"cSection"
Option Explicit

Private m_colSections As Collection

Public Enum SortOrderConstants
    soText
    soSize
End Enum

Public Function Add(ByVal Size As Double, ByVal Color As Long, ByVal Text As String, ByVal Key As String) As cSection

Dim oNewSection As cSection

    Set oNewSection = New cSection

    oNewSection.Color = Color
    oNewSection.Size = Size
    oNewSection.Text = Text
    oNewSection.Key = Key

    m_colSections.Add oNewSection, Key
    
    Set Add = oNewSection
    
    Set oNewSection = Nothing

End Function

Public Function Sort(eOrder As SortOrderConstants)

    Select Case eOrder
        Case soText: ' Only size for now.
        Case soSize
            If m_colSections.Count Then QuickSortD 1, m_colSections.Count
    End Select

End Function

Public Property Get Item(vIndexKey As Variant) As cSection
Attribute Item.VB_UserMemId = 0
  Set Item = m_colSections(vIndexKey)
End Property

Public Property Get Count() As Long
    Count = m_colSections.Count
End Property

Public Sub Remove(vIndexKey As Variant)
    m_colSections.Remove vIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_colSections.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set m_colSections = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_colSections = Nothing
End Sub

Private Sub QuickSortD(lo As Long, hi As Long)

' Sorts the Sections collection based on size.

Dim tc As Long, tk As String, ts As Double, tt As String
Dim tlo As Long, thi As Long, x As Double, y As Double

    tlo = lo: thi = hi

    x = m_colSections((lo + hi) / 2).Size

    While (tlo <= thi)
        While (m_colSections(tlo).Size < x And tlo < hi)
            tlo = tlo + 1
        Wend

        While (x < m_colSections(thi).Size And thi > lo)
           thi = thi - 1
        Wend

        If (tlo <= thi) Then
            tc = m_colSections(tlo).Color: tk = m_colSections(tlo).Key
            ts = m_colSections(tlo).Size: tt = m_colSections(tlo).Text
            
            m_colSections(tlo).Color = m_colSections(thi).Color
            m_colSections(tlo).Key = m_colSections(thi).Key
            m_colSections(tlo).Size = m_colSections(thi).Size
            m_colSections(tlo).Text = m_colSections(thi).Text
            
            m_colSections(thi).Color = tc: m_colSections(thi).Key = tk
            m_colSections(thi).Size = ts: m_colSections(thi).Text = tt

            tlo = tlo + 1
            thi = thi - 1
        End If
    Wend
    
   If (lo < thi) Then QuickSortD lo, thi
   If (tlo < hi) Then QuickSortD tlo, hi

End Sub

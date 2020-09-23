VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPieChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"cSections"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' API constants, types, and declarations.
Private Type tPoint
    x As Integer
    y As Integer
End Type

Private Declare Function Pie Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

' Non-API constants, types, and declarations.
Private Const TOTAL_RADIANS As Double = 6.28318530717959

Private Type SectionInfoEx
    Angle1 As Double
    Angle2 As Double
    Percent As Double
End Type

Private m_lColors(2) As Long
Private m_oSections  As cSections
Private m_oCanvas    As Object

Public Property Get Sections() As cSections
    If m_oSections Is Nothing Then Set m_oSections = New cSections
    Set Sections = m_oSections
End Property

Public Property Set Sections(n As cSections)
    Set m_oSections = n
End Property

Private Sub Class_Initialize()
' TODO: Make colors customizable.
    m_lColors(0) = vbRed
    m_lColors(1) = vbYellow
    m_lColors(2) = vbBlue
End Sub

Private Sub Class_Terminate()
    Set m_oSections = Nothing
End Sub

Public Property Set Canvas(n As Object)
    Set m_oCanvas = n
End Property

Public Property Get Canvas() As Object
    Set Canvas = m_oCanvas
End Property

Private Function pPolarToRect(ByVal dAngle As Double, ByRef iRadius As tPoint, _
                              ByRef pOrgin As tPoint) As tPoint
    pPolarToRect.x = iRadius.x * Cos(dAngle) + pOrgin.x
    pPolarToRect.y = iRadius.y * Sin(dAngle) + pOrgin.y
End Function

Public Function Draw()
   
Dim i           As Integer
Dim dTotalSize  As Double
Dim dPercent    As Double

Dim pOrgin As tPoint, pMid As tPoint
Dim pStart As tPoint, pStop As tPoint

Dim iMidRad As tPoint, iRad As tPoint   ' Not really "points", but 2 related integer values.
Dim iClrs As Integer                    ' # of colors to alterante between.

Dim SI() As SectionInfoEx

Dim cutoff_per As Double
Dim oStack As cSections
Dim stack_size As Double

    ReDim dPercents(1 To m_oSections.Count)
    ReDim SI(1 To m_oSections.Count)

    ' Get total size of all the sections.
    For i = 1 To m_oSections.Count
        dTotalSize = dTotalSize + m_oSections.Item(i).Size
    Next
    
    For i = 1 To m_oSections.Count
        dPercents(i) = m_oSections.Item(i).Size / dTotalSize
        'SI(i).Percent = m_oSections.Item(i).Size / dTotalSize
    Next

    fMain.List2.Clear
    
    pOrgin.x = m_oCanvas.ScaleWidth \ 2: pOrgin.y = m_oCanvas.ScaleHeight \ 2
    
    ' Change 0.7 to make pie more or less of Canvas.
    iRad.x = 0.7 * pOrgin.x: iRad.y = 0.7 * pOrgin.y
    
    m_oCanvas.Cls

' TODO: if Other doesn't reach say 10% by original method then add sections
' that are larger than cutoff to make Other=10% OR if Other only ends up
' containing one section, add as normal not as Other.

cutoff_per = 3 / 100 ' range 0 - 100

    ' Remove all sections that are smaller than cut-off percentage.
    i = m_oSections.Count
     
    Set oStack = New cSections
    Do While i
        If dPercents(i) < cutoff_per Then
            With m_oSections(i)
                oStack.Add .Size, .Color, .Text, .Key
                stack_size = stack_size + .Size
            End With
            m_oSections.Remove i
        End If
        i = i - 1
    Loop

    ' Add those all back as one section with text - "Other".
    If oStack.Count Then
        m_oSections.Add stack_size, vbGreen, "Other", ""
        pDrawStack oStack
    End If
    
    Set oStack = Nothing
    
    iClrs = IIf(m_oSections.Count Mod 2, 3, 2)
   
    If m_oSections.Count = 1 Then
        ' Draw circle
        m_oCanvas.FillColor = IIf(m_oSections.Item(1).Color = -1, _
                                  m_lColors(0), m_oSections.Item(1).Color)
        
        Ellipse m_oCanvas.hdc, pOrgin.x - iRad.x, pOrgin.y + iRad.y, _
                               pOrgin.x + iRad.x, pOrgin.y - iRad.y
    Else
    For i = 1 To m_oSections.Count
    
        dPercent = m_oSections.Item(i).Size / dTotalSize
    
        If i > 1 Then
            SI(i).Angle1 = SI(i - 1).Angle2 - 0.01
            SI(i).Angle2 = SI(i - 1).Angle2 + dPercent * TOTAL_RADIANS
        Else
            SI(i).Angle1 = SI(i).Angle2 - 0.01
            SI(i).Angle2 = SI(i).Angle2 + dPercent * TOTAL_RADIANS
        End If
        
        pStart = pPolarToRect(SI(i).Angle1, iRad, pOrgin)
        pStop = pPolarToRect(SI(i).Angle2, iRad, pOrgin)
        
        If m_oSections.Item(i).Color = -1 Then
            ' Auto-generate color.
            m_oCanvas.FillColor = m_lColors(i Mod iClrs)
        Else
            m_oCanvas.FillColor = m_oSections.Item(i).Color
        End If
    
        Pie m_oCanvas.hdc, pOrgin.x - iRad.x, pOrgin.y + iRad.y, _
                           pOrgin.x + iRad.x, pOrgin.y - iRad.y, _
                           pStop.x, pStop.y, pStart.x, pStart.y
                           
        fMain.List2.AddItem m_oSections.Item(i).Text & ":" & Format$(dPercent, "0.0%")
    Next
    End If
    
    ' Draw labels on top of all sections.
    For i = 1 To m_oSections.Count
        ' IDEA: Instead of mid, try to find closest angle to PI/2rad, 3PI/2rad.
        iMidRad.x = iRad.x / 1: iMidRad.y = iRad.y / 1
        pMid = pPolarToRect((SI(i).Angle1 + SI(i).Angle2) / 2, iMidRad, pOrgin)
        Call pDrawLabel(m_oSections.Item(i).Text, pMid.x, pMid.y)
    Next
    
End Function


Private Sub pDrawLabel(ByVal sCaption As String, ByVal x As Integer, ByVal y As Integer)
    
Dim iLabel As tPoint, t As tPoint

    iLabel.x = m_oCanvas.TextWidth(sCaption)
    iLabel.y = m_oCanvas.TextHeight(sCaption)
    
    ' TODO: Print labels on top of all sections.
    t.x = x - iLabel.x \ 2: t.y = y - iLabel.y \ 2
    
    ' Draw outlined box.
    m_oCanvas.Line (t.x - 3, t.y - 3)-Step(iLabel.x + 6, iLabel.y + 6), vbWindowText, B
    m_oCanvas.Line (t.x - 2, t.y - 2)-Step(iLabel.x + 4, iLabel.y + 4), vb3DHighlight, BF
    
    ' Draw caption.
    m_oCanvas.CurrentX = t.x: m_oCanvas.CurrentY = t.y
    m_oCanvas.Print sCaption
    
End Sub

Private Sub pDrawStack(o As cSections)

Dim i As Integer

    For i = 1 To o.Count
        fMain.List2.AddItem "Other - " & o(i).Text
    Next
'
'' Draw a stacked column chart.

'Dim dTotal As Double
'Dim iColH As Integer
'Dim iSecH As Integer
'Dim startY As Integer
'Dim dPercent As Double
'Dim dPercents() As Double
'Dim iSecHs() As Integer
'Dim iTxtH As Integer
'Dim dBoost As Double
'
'    iColH = fMain.picLegend.ScaleHeight
'
'    ' Smallest a section of the stack can be.
'    iTxtH = fMain.picLegend.TextHeight("X")
'
'    ' Get total size.
'    For i = 1 To o.Count
'        dTotal = dTotal + o(i).Size
'    Next
'
'    ReDim dPercents(1 To o.Count)
'    ReDim iSecHs(1 To o.Count)
'
'    ' Find out smallest section size and what it takes to make it larger.
'    For i = 1 To o.Count
'
'        iSecH = (o(i).Size / dTotal) * iColH
'
'        ' If section height is less than required minimum..
'        If iSecH < iTxtH Then
'            dBoost = iTxtH / iSecH
'            Exit For
'        End If
'
'    Next
'
'    ' Assumes o is sorted by Size.
'    For i = 1 To o.Count
'        'dPercents(i) = o(i).Size / dTotal
'        iSecHs(i) = (o(i).Size / dTotal) * iColH * dBoost
'    Next
'
'    fMain.picLegend.Cls
'
'    For i = 1 To o.Count
'        'dPercent = o(i).Size / dTotal
'        dPercent = iSecHs(i) / iColH
'
'        iSecH = dPercent * iColH
'
'        fMain.picLegend.Line (0, startY)-Step(10, iSecH), m_lColors(i Mod 2), BF
'        fMain.picLegend.CurrentX = 12
'        fMain.picLegend.CurrentY = startY
'        fMain.picLegend.Print o(i).Text
'        startY = startY + iSecH
'    Next
'
End Sub


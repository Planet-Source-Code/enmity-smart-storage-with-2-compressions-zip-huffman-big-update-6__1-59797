VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type FILTERSTRUCT
    Description As String
    Extensions As String
End Type

Private m_Filters() As FILTERSTRUCT

Public Sub Clear()
    Erase m_Filters
End Sub

Public Property Get Filter(ByVal vData As Long) As String
    If Not Count Then Exit Property
    If UBound(m_Filters) < vData Then Exit Property
        
    Filter = m_Filters(vData).Description + ";" + m_Filters(vData).Extensions
End Property

Public Property Get Count() As Long

    Dim i As Long
    
    On Error Resume Next
    i = -1&
    
    i = UBound(m_Filters)
    Count = i + 1

End Property

Public Sub AddFilter(ByVal szDescription As String, ByVal szFilter As String)

    On Error Resume Next
    Dim i As Long
    
    i = -1&
    i = UBound(m_Filters)
    
    i = i + 1
    
    ReDim Preserve m_Filters(0 To i)
    
    With m_Filters(i)
        .Description = szDescription
        .Extensions = szFilter
    End With
    
End Sub

Public Sub Remove(ByVal Index As Long)
    Dim xFilters() As FILTERSTRUCT
    
    Dim i As Long, _
        j As Long
    
    On Error Resume Next
    
    i = -1&
    i = UBound(m_Filters)

    If i = -1& Then Exit Sub
    
    ReDim xFilters(0 To i)
    
    j = 0
    For i = 0 To UBound(m_Filters)
    
        If i <> Index Then
            
            CopyMemory xFilters(j), m_Filters(i), Len(m_Filters(i))
            j = j + 1
        
        End If
    
    Next i

    j = j - 1
    
    If j < 0 Then
        Erase m_Filters
        Exit Sub
    End If
    
    ReDim Preserve xFilters(0 To j)
    
    CopyMemory m_Filters(0), xFilters(0), Len(m_Filters(0)) * (j + 1)
    Erase xFilters
    
End Sub

Friend Property Get OfnFilter() As String
    Dim tStr As String
    
    Dim i As Long
    
    For i = 0 To (Count - 1)

        With m_Filters(i)
            tStr = tStr + .Description + Chr(0) + .Extensions + Chr(0)
        End With

    Next i
    tStr = tStr + Chr(0)
    OfnFilter = tStr
    
End Property

Attribute VB_Name = "mListViewEx"
'========================================================================================
' Module:        mListViewEx.bas (.Sort() routines)
' Last revision: 2004.09.04
'========================================================================================

Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type LVITEM_lp
    mask       As Long
    iItem      As Long
    iSubItem   As Long
    State      As Long
    stateMask  As Long
    pszText    As Long
    cchTextMax As Long
    iImage     As Long
    lParam     As Long
    iIndent    As Long
End Type

Private Type LVFINDINFO
    flags       As Long
    psz         As Long
    lParam      As Long
    pt          As POINTAPI
    vkDirection As Long
End Type
 
Private Const LVFI_PARAM      As Long = &H1
Private Const LVIF_TEXT       As Long = &H1

Private Const LVM_FIRST       As Long = &H1000
Private Const LVM_FINDITEM    As Long = (LVM_FIRST + 13)
Private Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)
Private Const LVM_SORTITEMS   As Long = (LVM_FIRST + 48)
     
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'//

Private m_uLVFI      As LVFINDINFO
Private m_uLVI       As LVITEM_lp
Private m_lColumn    As Long

Private m_PRECEDE    As Long
Private m_FOLLOW     As Long

'//

Private Function pvCompareIndex( _
                 ByVal lParam1 As Long, _
                 ByVal lParam2 As Long, _
                 ByVal hWnd As Long) As Long

    If (lParam1 > lParam2) Then
        pvCompareIndex = m_PRECEDE
    ElseIf (lParam1 < lParam2) Then
        pvCompareIndex = m_FOLLOW
    End If
End Function

Private Function pvCompareText( _
                 ByVal lParam1 As Long, _
                 ByVal lParam2 As Long, _
                 ByVal hWnd As Long) As Long

  Dim val1 As String
  Dim val2 As String
     
    val1 = LCase$(pvGetItemText(hWnd, lParam1))
    val2 = LCase$(pvGetItemText(hWnd, lParam2))
     
    If (val1 > val2) Then
        pvCompareText = m_PRECEDE
    ElseIf (val1 < val2) Then
        pvCompareText = m_FOLLOW
    End If
End Function

Private Function pvCompareTextSensitive( _
                 ByVal lParam1 As Long, _
                 ByVal lParam2 As Long, _
                 ByVal hWnd As Long) As Long

  Dim val1 As String
  Dim val2 As String
     
    val1 = pvGetItemText(hWnd, lParam1)
    val2 = pvGetItemText(hWnd, lParam2)
     
    If (val1 > val2) Then
        pvCompareTextSensitive = m_PRECEDE
    ElseIf (val1 < val2) Then
        pvCompareTextSensitive = m_FOLLOW
    End If
End Function

Private Function pvCompareValue( _
                 ByVal lParam1 As Long, _
                 ByVal lParam2 As Long, _
                 ByVal hWnd As Long) As Long

  Dim val1 As Double
  Dim val2 As Double
     
    val1 = pvGetItemValue(hWnd, lParam1)
    val2 = pvGetItemValue(hWnd, lParam2)
     
    If (val1 > val2) Then
        pvCompareValue = m_PRECEDE
    ElseIf (val1 < val2) Then
        pvCompareValue = m_FOLLOW
    End If
End Function

Private Function pvCompareDate( _
                 ByVal lParam1 As Long, _
                 ByVal lParam2 As Long, _
                 ByVal hWnd As Long) As Long

  Dim val1 As Date
  Dim val2 As Date
     
    val1 = pvGetItemDate(hWnd, lParam1)
    val2 = pvGetItemDate(hWnd, lParam2)
     
    If (val1 > val2) Then
        pvCompareDate = m_PRECEDE
    ElseIf (val1 < val2) Then
        pvCompareDate = m_FOLLOW
    End If
End Function

'//

Private Function pvGetItemText( _
                 hWnd As Long, _
                 lParam As Long) As String
  
  Dim lIdx   As Long
  Dim a(261) As Byte
  Dim lLen   As Long
    
    With m_uLVFI
        .flags = LVFI_PARAM
        .lParam = lParam
    End With
    lIdx = SendMessage(hWnd, LVM_FINDITEM, -1, m_uLVFI)

    With m_uLVI
        .mask = LVIF_TEXT
        .pszText = VarPtr(a(0))
        .cchTextMax = UBound(a)
        .iSubItem = m_lColumn
    End With
    lLen = SendMessage(hWnd, LVM_GETITEMTEXT, lIdx, m_uLVI)
    
    pvGetItemText = Left$(StrConv(a(), vbUnicode), lLen)
End Function

Private Function pvGetItemValue( _
                 hWnd As Long, _
                 lParam As Long) As Double
  
  Dim lIdx   As Long
  Dim a(261) As Byte
  Dim lLen   As Long

    With m_uLVFI
        .flags = LVFI_PARAM
        .lParam = lParam
    End With
    lIdx = SendMessage(hWnd, LVM_FINDITEM, -1, m_uLVFI)

    With m_uLVI
        .mask = LVIF_TEXT
        .pszText = VarPtr(a(0))
        .cchTextMax = UBound(a)
        .iSubItem = m_lColumn
    End With
    lLen = SendMessage(hWnd, LVM_GETITEMTEXT, lIdx, m_uLVI)
    
    If (lLen) Then
        pvGetItemValue = CDbl(Left$(StrConv(a(), vbUnicode), lLen))
      Else
        pvGetItemValue = -1
    End If
End Function

Private Function pvGetItemDate( _
                 hWnd As Long, _
                 lParam As Long) As Date
  
  Dim lIdx   As Long
  Dim a(261) As Byte
  Dim sText  As String
  Dim lLen   As Long
    
    With m_uLVFI
        .flags = LVFI_PARAM
        .lParam = lParam
    End With
    lIdx = SendMessage(hWnd, LVM_FINDITEM, -1, m_uLVFI)
     
    With m_uLVI
        .mask = LVIF_TEXT
        .pszText = VarPtr(a(0))
        .cchTextMax = UBound(a)
        .iSubItem = m_lColumn
    End With
    lLen = SendMessage(hWnd, LVM_GETITEMTEXT, lIdx, m_uLVI)
    
    sText = Left$(StrConv(a(), vbUnicode), lLen)
    If (IsDate(sText)) Then
        pvGetItemDate = sText
    End If
End Function

Private Function AddressOfFunction(lpfn As Long) As Long
    AddressOfFunction = lpfn
End Function

'//

Public Function Sort( _
                ByVal hListView As Long, _
                ByVal Column As Integer, _
                ByVal SortOrder As eSortOrderConstants, _
                ByVal SortType As eSortTypeConstants _
                ) As Boolean

  Dim lRet As Long
  
    m_lColumn = CLng(Column)
        
    Select Case SortOrder
        
        Case [soDefault]
            
            m_PRECEDE = 1
            m_FOLLOW = -1
            lRet = SendMessageLong(hListView, LVM_SORTITEMS, hListView, AddressOfFunction(AddressOf pvCompareIndex))
            
        Case [soAscending], [soDescending]
        
            m_PRECEDE = SortOrder
            m_FOLLOW = -SortOrder
            
            Select Case SortType
                Case [stString]
                    lRet = SendMessageLong(hListView, LVM_SORTITEMS, hListView, AddressOfFunction(AddressOf pvCompareText))
                Case [stStringSensitive]
                    lRet = SendMessageLong(hListView, LVM_SORTITEMS, hListView, AddressOfFunction(AddressOf pvCompareTextSensitive))
                Case [stNumeric]
                    lRet = SendMessageLong(hListView, LVM_SORTITEMS, hListView, AddressOfFunction(AddressOf pvCompareValue))
                Case [stDate]
                    lRet = SendMessageLong(hListView, LVM_SORTITEMS, hListView, AddressOfFunction(AddressOf pvCompareDate))
            End Select
    End Select
    
    Sort = CBool(lRet)
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum OpenFileDlgFlags
    
    ofnAllowMultiselect = &H200
    ofnCreatePrompt = &H2000
'    ofnEnableHook = &H20
'    ofnEnabledTemplate = &H40
'    ofnEnableTemplateHandle = &H80
    ofnExplorer = &H80000
    ofnExtensionDifferent = &H400
    ofnFileMustExist = &H1000
    ofnHideReadOnly = &H4
    ofnLongNames = &H200000
    ofnNoChangeDir = &H8
    ofnNoDereferenceLinks = &H100000
    ofnNoLongNames = &H40000
    ofnNoNetworkButton = &H20000
    ofnNoReadOnlyReturn = &H8000
    ofnNoTestFileCreate = &H10000
    ofnNoValidate = &H100
    ofnOverwritePrompt = &H2
    ofnPathMustExist = &H800
    ofnReadOnly = &H1
    ofnShareAware = &H4000
    ofnShareFallThrough = 2
    ofnShareNoWarn = 1
    ofnShareWarn = 0
    ofnShowHelp = &H10

End Enum

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
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
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private m_ofn As OPENFILENAME

Private m_InitDir As String
Private m_Filename As String
Private m_FileTitle As String

Private m_Flags As OpenFileDlgFlags
Private m_FilterIndex As Long
Private m_Title As String
Private m_CustomFilter As String
Private m_DefExt As String
Private m_Cancel As Boolean

Private m_Filters As cFilters

Public Property Get Cancelled() As Boolean
    Cancelled = m_Cancel
End Property

Public Property Get Filters() As cFilters
    Set Filters = m_Filters
End Property

Public Property Let InitDir(ByVal vData As String)
    m_InitDir = vData
End Property

Public Property Get InitDir() As String
    InitDir = m_InitDir
End Property

Public Property Let DefExt(ByVal vData As String)
    m_DefExt = vData
End Property

Public Property Get DefExt() As String
    DefExt = m_DefExt
End Property

Public Property Let Title(ByVal vData As String)
    m_Title = vData
End Property

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let FilterIndex(ByVal vData As Long)
    m_FilterIndex = vData
End Property

Public Property Get FilterIndex() As Long
    FilterIndex = m_FilterIndex
End Property

Public Property Let Flags(ByVal vData As OpenFileDlgFlags)
    m_Flags = vData
End Property

Public Property Get Flags() As OpenFileDlgFlags
    Flags = m_Flags
End Property

Public Property Let FileName(ByVal vData As String)
    m_Filename = vData
End Property

Public Property Get FileName() As String
    FileName = m_Filename
End Property

Public Property Let FileTitle(ByVal vData As String)
    m_FileTitle = vData
End Property

Public Property Get FileTitle() As String
    FileTitle = m_FileTitle
End Property

Public Function ShowOpen(ByVal hwnd As Long) As Boolean
    
    m_ofn.hWndOwner = hwnd
    m_ofn.Flags = m_Flags
    m_ofn.lpstrDefExt = m_DefExt
    m_ofn.lpstrFile = m_Filename + String(255 - Len(m_Filename), Chr(0)) + Chr(0)
    m_ofn.nMaxFile = 255
    m_ofn.lpstrFileTitle = m_FileTitle + String(255 - Len(m_Filename), Chr(0)) + Chr(0)
    m_ofn.nMaxFileTitle = 255
    m_ofn.nMaxCustFilter = 255
    m_ofn.lpstrCustomFilter = m_CustomFilter + String(255 - Len(m_CustomFilter), Chr(0)) + Chr(0)
    m_ofn.lpstrFilter = m_Filters.OfnFilter
    m_ofn.lpstrInitialDir = m_InitDir + Chr(0)
    m_ofn.nFilterIndex = m_FilterIndex
    m_ofn.lpstrTitle = m_Title + Chr(0)
    
    ShowOpen = CBool(GetOpenFileName(m_ofn))
    If Not ShowOpen Then m_Cancel = True
    
    m_Filename = ClipNulls(m_ofn.lpstrFile)
    m_FileTitle = ClipNulls(m_ofn.lpstrFileTitle)
    m_FilterIndex = m_ofn.nFilterIndex
    
End Function

Public Function ShowSave(ByVal hwnd As Long) As Boolean
    Dim i As Long
    
    m_ofn.hWndOwner = hwnd
    m_ofn.Flags = m_Flags
    m_ofn.lpstrDefExt = m_DefExt
    m_ofn.lpstrFile = m_Filename + String(255 - Len(m_Filename), Chr(0)) + Chr(0)
    m_ofn.nMaxFile = 255
    m_ofn.lpstrFileTitle = m_FileTitle + String(255 - Len(m_Filename), Chr(0)) + Chr(0)
    m_ofn.nMaxFileTitle = 255
    m_ofn.nMaxCustFilter = 255
    m_ofn.lpstrCustomFilter = m_CustomFilter + String(255 - Len(m_CustomFilter), Chr(0)) + Chr(0)
    m_ofn.lpstrFilter = m_Filters.OfnFilter
    m_ofn.lpstrInitialDir = m_InitDir + Chr(0)
    m_ofn.nFilterIndex = m_FilterIndex
    m_ofn.lpstrTitle = m_Title + Chr(0)
    
    ShowSave = CBool(GetSaveFileName(m_ofn))
    If Not ShowSave Then m_Cancel = True
    
    m_Filename = ClipNulls(m_ofn.lpstrFile)
    m_FileTitle = ClipNulls(m_ofn.lpstrFileTitle)
    m_FilterIndex = m_ofn.nFilterIndex
    
End Function

Private Sub Class_Initialize()
    
    m_ofn.lStructSize = Len(m_ofn)
    m_ofn.hInstance = App.hInstance
    m_ofn.lpstrInitialDir = App.Path
    
    m_Flags = ofnExplorer + ofnFileMustExist + ofnLongNames + ofnShareAware
    Set m_Filters = New cFilters
    m_Filters.AddFilter "All Files (*.*)", "*.*"
    
End Sub

Private Function ClipNulls(ByVal vString As String) As String

    If InStr(1, vString, Chr(0)) <> 0 Then
        ClipNulls = Mid(vString, 1, InStr(1, vString, Chr(0)) - 1)
    Else
        ClipNulls = vString
    End If
End Function


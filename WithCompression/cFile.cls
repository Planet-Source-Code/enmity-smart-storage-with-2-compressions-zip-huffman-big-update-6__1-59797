VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MAX_PATH = 260

Private Const INVALID_HANDLE_VALUE = -1

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Private Declare Function FindClose Lib "kernel32" _
  (ByVal hFindFile As Long) As Long

Private Declare Function FindFirstFile Lib "kernel32" _
   Alias "FindFirstFileA" _
  (ByVal lpFileName As String, _
   lpFindFileData As WIN32_FIND_DATA) As Long

Private Const FILE_ATTRIBUTE_DIRECTORY = &H10



Public Function DoesFileExistEx(ByVal sFile As String) As Boolean
    
    Dim FindData As WIN32_FIND_DATA
    Dim GetFile As Long
       
    GetFile = FindFirstFile(sFile, FindData)
        
    If GetFile <> INVALID_HANDLE_VALUE Then
        If (FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
            DoesFileExistEx = False '??
        Else
            DoesFileExistEx = True
        End If
    Else
        DoesFileExistEx = False
    End If
    
    FindClose GetFile
 
End Function


Public Function GetFileBaseName(ByVal strFileName As String) As String
        If Len(strFileName) > 2 Then 'x:\a.bbb
            If InStrRev(strFileName, "\") > 0 Then
                strFileName = Right$(strFileName, Len(strFileName) - InStrRev(strFileName, "\"))
            Else
            End If
            
            If InStr(strFileName, ".") <> 0 Then
                GetFileBaseName = Left$(strFileName, InStrRev(strFileName, ".") - 1)
            Else
                GetFileBaseName = strFileName
            End If
        Else
            GetFileBaseName = ""
        End If
End Function


Public Function GetFilePath(ByVal strFile As String) As String
        
        GetFilePath = Left$(strFile, InStrRev(strFile, "\") - 1)
        
End Function


Public Function GetFileName(ByVal strFile As String) As String
        
        GetFileName = Right$(strFile, Len(strFile) - InStrRev(strFile, "\"))

End Function


Public Function GetFileExtName(ByVal strFileName As String) As String
        
        Dim o_strResult As String
        
        o_strResult = Right$(strFileName, Len(strFileName) - InStrRev(strFileName, "\"))
        If InStrRev(o_strResult, ".") <> 0 Then
            o_strResult = Right$(o_strResult, Len(o_strResult) - InStrRev(o_strResult, "."))
        Else
            o_strResult = ""
        End If
        
        GetFileExtName = o_strResult
    
End Function


Public Function LoadTextFromFile(ByVal strFile As String) As String
        Dim o_intFileNum As Integer
        Dim o_strContent As String
        
        o_intFileNum = FreeFile
        
        Open strFile For Binary As #o_intFileNum
        
        o_strContent = Space(LOF(o_intFileNum))
        
        Get #o_intFileNum, , o_strContent

        Close #o_intFileNum
        
        LoadTextFromFile = o_strContent
End Function


Public Function LoadBytesFromFile(ByVal strFile As String, _
                                  ByRef bytContent() As Byte, _
                                  Optional ByVal lngStart As Long = 1, _
                                  Optional ByVal lngCount As Long = 0) As Boolean
    
    Dim o_intFileNum As Integer
    Dim o_bytArray() As Byte

    If (DoesFileExistEx(strFile)) Then
        o_intFileNum = FreeFile
        
        Open strFile For Binary As #o_intFileNum
        
        If lngCount <= 0 Then
            lngCount = LOF(o_intFileNum) + lngCount
        Else
        End If
        
        If lngStart <= 0 Then
            lngStart = LOF(o_intFileNum) + lngStart + 1
        Else
        End If
        
        ReDim o_bytArray(0 To lngCount - 1)
        
        Get #o_intFileNum, lngStart, o_bytArray()
        
        Close #o_intFileNum
        
        LoadBytesFromFile = o_bytArray
    Else
    End If
    
End Function

Public Function SaveContentToFile(ByVal strFile As String, ByRef bytContent() As Byte) As Boolean
        Dim o_intFileNum As Integer
        
        o_intFileNum = FreeFile
        
        Open strFile For Binary As #o_intFileNum

        Put #o_intFileNum, , bytContent

        Close #o_intFileNum
        
        SaveContentToFile = True
End Function


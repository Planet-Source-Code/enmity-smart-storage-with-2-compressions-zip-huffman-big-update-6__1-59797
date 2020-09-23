VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Smart Storage with 2 compressions(zip/huffman)
'
'Smart Storage is something like a file packager, you can add any file(s) of
'any format to the storage file(AddUpdateItem function) and remove them
'(DeleteItem function), of course.
'
'You can choose zlib/huffman or no compression for each file.
'
'It uses index(file) technique for a better performance, thus it has
'VaccumStorage()function.
'
'And, the most inportant part is that it uses serialized section(chunk) technique
'to handle the CRC & compression of large file.
'
'And this project is originally aimed to be the storage part of Carles P.V's
'Thumbnailer 1.0 (image thumbnailer-viewer with GDI+)(http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=59677&lngWId=1), thus it includes the PictureFromByteStream() function to directly return a picture object from a byte array!
'
'Beside that, there are: GetItemText() function for direct return of plain text
'content; GetItemPicture() function for direct return of picture object according
'to the index; SaveItemToFile() function to save the content to disk file according
'to the index.
'
'Some of the codes are not written by me, such as cHuffman, cCRC. They are from psc,
'now they are back for you guys:)
'
'The zlib dll is generated at runtime of sample test since psc will remove all PE
'file. It is easy for you to remove the zlib dependancy, that's why I only enclosed
'it in the sample project.
'
'And I make all these functions into a class for handy usage. I also use Copymemory
'for a better performance. Please feel free to leave any comments, bugs or suggestions are welcome!
'
'Usage:
'1.Simply open the sample storage, then click listed files, content displayed,
'  doubleclick to save.
'2.Select DB File->Create->Open->Select Item File, add it...
'
'Sorry for lack of code comments, but I think that the method is really simple.

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

'default chunk size, bigger chunk for better performance but more memory usage
Private Const mc_lngChunkSize As Long = 1024 '1048576 '1024*1024 = 1M, a bigger chunk reduces index!
Private Const mc_lngIndexHeaderSignatureSize As Long = 6
Private Const mc_lngIndexHeaderVersionSize As Long = 2
Private Const mc_lngIndexHeaderSize As Long = mc_lngIndexHeaderSignatureSize + mc_lngIndexHeaderVersionSize
Private Const mc_strVersion As String = "06"

Public Enum enumCompresssion
    Normal = 0
    Zlib = 1
    Huffman = 2
End Enum

Public Enum enumStorageType
    File = 0
    ByteArray = 1
    Text = 2
    Image = 3
End Enum

Public Enum enumFilterType
    Key = 1
    [Type] = 2
    ItemData = 4
    CreatedDate = 8
    LastModifiedDate = 16
    Memo = 32
    Category = 64
    Flag = 128
End Enum

Public Enum enumOpenStorageResult
    OK = 0
    VersionTooLow = 1
    VersionTooHigh = 2
    Mailformed = 3
    StorageNotFound = 4
    Error = 5
End Enum

Public Enum enumReadChunkResult
    OK = 0
    OffsetOutOfBound = 1
    SizeToLarge = 2
    StorageNotReady = 3
End Enum

Public Enum enumAddUpdateItemResult
    OK = 0
    FileNotFound = 1
    StorageNotReady = 2
    Error = 3
End Enum

Public Type typeCompressedSection
    lngOriginalSize As Long
    lngCompressedSize As Long
End Type

Public Type typeIndexRecord
    strKey As String
    udeType As Long
    strType As String
    lngStart As Long
    lngItemData As Long
    udtCompressedSection() As typeCompressedSection
    dteCreatedDate As Date
    dteLastModifiedDate As Date
    strCategory As String
    strMemo As String
    udeFlag As Long
    strCRC As String
End Type

Private m_strVersion As String
Private m_colIndex As Collection
Private m_intIndexFileNum As Integer
Private m_intStorageFileNum As Integer
Private m_lngChunkSize As Long
Private m_bytIndexHeader(5) As Byte
Private m_bytSplitter(5) As Byte
Private m_strStorageFilePath As String
Private m_blnStorageReady As Boolean
Private m_blnUpdated As Boolean
Private m_ucmFile As New cFile
Private m_ucmCompress As New cCompression
Private m_ucmHuffman As New cHuffman
Private m_ucmCRC As New cCRC
Private m_ucmDIBs As New cDIBsEx



Public Property Get StorageFilePath() As String
        StorageFilePath = m_strStorageFilePath
End Property

Public Property Let StorageFilePath(ByVal strStorageFilePath As String)
        m_strStorageFilePath = strStorageFilePath
End Property


Public Property Get Version() As String
        Version = m_strVersion
End Property

Public Property Let Version(ByVal strVersion As String)
        m_strVersion = strVersion
End Property


Public Property Get ChunkSize() As Long
        ChunkSize = m_lngChunkSize
End Property

Public Property Let ChunkSize(ByVal lngChunkSize As Long)
        m_lngChunkSize = lngChunkSize
End Property


Public Property Get StorageReady() As Boolean
        StorageReady = m_blnStorageReady
End Property

Public Property Let StorageReady(ByVal blnStorageReady As Boolean)
        m_blnStorageReady = blnStorageReady
End Property


Public Property Get Items() As Collection
        Set Items = m_colIndex
End Property


Public Function GetStorageTypeName(ByVal udeType As enumStorageType) As String
        Select Case udeType
            Case enumStorageType.File
                GetStorageTypeName = "File"
            Case enumStorageType.Image
                GetStorageTypeName = "Image"
            Case enumStorageType.ByteArray
                GetStorageTypeName = "ByteArray"
            Case enumStorageType.Text
                GetStorageTypeName = "Text"
        End Select
End Function


Public Function GetCompressionTypeName(ByVal udeType As enumCompresssion) As String
        Select Case udeType
            Case enumCompresssion.Normal
                GetCompressionTypeName = "Normal"
            Case enumCompresssion.Zlib
                GetCompressionTypeName = "Zip"
            Case enumCompresssion.Huffman
                GetCompressionTypeName = "Huffman"
        End Select
End Function


Public Function GetOriginalSize(ByVal lngItemIndex As Long) As Long
        Dim o_lngItems As Long
        Dim o_lngResult As Long
        Dim o_udtIndexRecord As typeIndexRecord
        
        o_udtIndexRecord = m_colIndex.Item(lngItemIndex)
        With o_udtIndexRecord
            For o_lngItems = LBound(.udtCompressedSection) To UBound(.udtCompressedSection)
                o_lngResult = o_lngResult + .udtCompressedSection(o_lngItems).lngOriginalSize
            Next
        End With
        
        GetOriginalSize = o_lngResult
End Function


Public Function GetCompressedSize(ByVal lngItemIndex As Long) As Long
        Dim o_lngItems As Long
        Dim o_lngResult As Long
        Dim o_udtIndexRecord As typeIndexRecord
        
        o_udtIndexRecord = m_colIndex.Item(lngItemIndex)
        With o_udtIndexRecord
            For o_lngItems = LBound(.udtCompressedSection) To UBound(.udtCompressedSection)
                o_lngResult = o_lngResult + .udtCompressedSection(o_lngItems).lngCompressedSize
            Next
        End With
        
        GetCompressedSize = o_lngResult
End Function


'get items according to filters
Public Function FilterItems(ByVal strFilter As String, ByVal udeFilter As enumFilterType) As Collection
        If m_blnStorageReady Then
            Dim o_lngItems As Long
            Dim o_lngCount As Long
            Dim o_colFilteredItems As New Collection
            Dim o_blnResult As Boolean
            
            With m_colIndex
                For o_lngItems = 1 To .Count
                    o_blnResult = False
                    
                    If udeFilter And Key Then
                        o_blnResult = .Item(o_lngItems).strKey Like "*" & strFilter & "*"
                    End If
                    If (Not o_blnResult) And (udeFilter And CreatedDate) Then
                        o_blnResult = .Item(o_lngItems).dteCreatedDate Like "*" & strFilter & "*"
                    End If
                    If (Not o_blnResult) And (udeFilter And LastModifiedDate) Then
                        o_blnResult = .Item(o_lngItems).dteLastModifiedDate Like "*" & strFilter & "*"
                    End If
                    If (Not o_blnResult) And (udeFilter And ItemData) Then
                        o_blnResult = .Item(o_lngItems).lngItemData = Val(strFilter)
                    End If
                    If (Not o_blnResult) And (udeFilter And [Type]) Then
                        o_blnResult = .Item(o_lngItems).strType Like "*" & strFilter & "*"
                    End If
                    If (Not o_blnResult) And (udeFilter And Flag) Then
                        o_blnResult = .Item(o_lngItems).udeFlag = Val(strFilter)
                    End If
                    If (Not o_blnResult) And (udeFilter And Memo) Then
                        o_blnResult = .Item(o_lngItems).strMemo Like "*" & strFilter & "*"
                    End If
                    If (Not o_blnResult) And (udeFilter And Category) Then
                        o_blnResult = .Item(o_lngItems).strCategory Like "*" & strFilter & "*"
                    End If
                    
                    If o_blnResult Then
                        o_lngCount = o_lngCount + 1
                        o_colFilteredItems.Add .Item(o_lngItems), "Index" & o_lngCount
                    End If
                Next
            End With
            
            Set FilterItems = o_colFilteredItems
        Else
        End If
End Function


'quick & dirty way to find a byte array in another byte array, hey, it just WORKS!
Private Function FindSection(ByRef bytChunk() As Byte, ByRef bytHeader() As Byte, ByVal lngPos As Long) As Long
        FindSection = InStrB(lngPos, bytChunk, bytHeader, vbBinaryCompare)
End Function


'create the storage, simply add the header
Public Function CreateStorage() As Boolean

        Dim o_intFileNum As Integer
        Dim o_blnResult As Boolean
        
        o_intFileNum = FreeFile
        
        Open GetIndexFile() For Binary As #o_intFileNum
        
        WriteHeader o_intFileNum
        
        Close #o_intFileNum
        
        o_blnResult = True
        
        CreateStorage = o_blnResult
        
End Function


'simply wrtie the header with signature & version
Private Sub WriteHeader(ByVal intFileNum As Integer)
        Dim o_bytContent() As Byte
        
        o_bytContent = StrConv(mc_strVersion, vbFromUnicode)
        
        Put #intFileNum, , m_bytIndexHeader
        
        Put #intFileNum, , o_bytContent
        
        Erase o_bytContent
End Sub


Private Function GetIndexFile() As String
        GetIndexFile = m_ucmFile.GetFilePath(m_strStorageFilePath) & "\" & m_ucmFile.GetFileBaseName(m_strStorageFilePath) & ".ind"
End Function


'open the storage, simply reads the index file
Public Function OpenStorage() As enumOpenStorageResult
    
    On Error GoTo HandleError
    
        Dim o_udeResult As enumOpenStorageResult
        
        If m_ucmFile.DoesFileExistEx(GetIndexFile()) Then
            Dim o_lngSize As Long
            Dim o_udtIndexRecord As typeIndexRecord
            Dim o_lngCount As Long
            
            Set m_colIndex = New Collection
            
            m_intIndexFileNum = FreeFile
            
            Open GetIndexFile For Binary As #m_intIndexFileNum
            
            o_lngSize = LOF(m_intIndexFileNum)
            
            If o_lngSize >= mc_lngIndexHeaderSize Then
                Dim o_bytChunk() As Byte
                ReDim o_bytChunk(0 To mc_lngIndexHeaderSignatureSize - 1) As Byte
                Get #m_intIndexFileNum, , o_bytChunk
                
                'verify the header
                If FindSection(o_bytChunk, m_bytIndexHeader, 1) Then
                    ReDim o_bytChunk(0 To 1) As Byte
                    Get #m_intIndexFileNum, , o_bytChunk
                    
                    m_strVersion = StrConv(o_bytChunk, vbUnicode)
                    
                    If Val(m_strVersion) = Val(mc_strVersion) Then
                        If o_lngSize > mc_lngIndexHeaderSize Then
                            Seek #m_intIndexFileNum, mc_lngIndexHeaderSize + 1
                            
                            Do Until EOF(m_intIndexFileNum)
                                Get #m_intIndexFileNum, , o_udtIndexRecord
                                
                                If o_udtIndexRecord.strKey <> "" Then
                                    o_lngCount = o_lngCount + 1
                                    
                                    m_colIndex.Add o_udtIndexRecord, "Index" & o_lngCount
                                Else
                                    Exit Do
                                End If
                            Loop
                        End If
                    
                        m_intStorageFileNum = FreeFile
                        
                        Open m_strStorageFilePath For Binary As #m_intStorageFileNum
                        
                        m_blnStorageReady = True
                        
                        m_blnUpdated = True
                        
                        o_udeResult = enumOpenStorageResult.OK
                    ElseIf Val(m_strVersion) > Val(mc_strVersion) Then
                        o_udeResult = VersionTooHigh
                    Else
                        o_udeResult = VersionTooLow
                    End If
                Else
                    o_udeResult = Mailformed
                End If
            Else
                o_udeResult = Mailformed
            End If
        Else
            o_udeResult = StorageNotFound
        End If
        
        If o_udeResult <> enumOpenStorageResult.OK Then
            Close #m_intIndexFileNum
        End If
        
        OpenStorage = o_udeResult
    
    Exit Function
    
HandleError:

    Close #m_intIndexFileNum

    OpenStorage = enumOpenStorageResult.Error
        
End Function


'close both index & Storage file
Public Function CloseStorage() As Boolean
        Dim o_blnResult As Boolean

        If m_blnStorageReady Then
            Close #m_intIndexFileNum
            Close #m_intStorageFileNum
            
            o_blnResult = True
                        
            m_blnStorageReady = False
        Else
            o_blnResult = True
        End If
        
        m_blnUpdated = True
        
        CloseStorage = o_blnResult
        
End Function


'find the index according to the specified file path
Public Function GetItemIndex(ByVal strKey As String) As Long
        Dim o_lngResult As Long
            
        o_lngResult = 0
            
        If m_blnStorageReady Then
            Dim o_lngItems As Long
            
            With m_colIndex
                For o_lngItems = 1 To .Count
                    With .Item(o_lngItems)
                        If LCase(.strKey) = LCase(strKey) Then
                            o_lngResult = o_lngItems
                            Exit For
                        End If
                    End With
                Next
            End With
        End If
        
        GetItemIndex = o_lngResult
End Function


'get the actual content according to the index
Public Function GetItemContent(ByVal lngItemIndex As Long, ByRef bytContent() As Byte, Optional ByVal blnDecode As Boolean = True) As Boolean
        Dim o_blnResult As Boolean
        
        If m_blnStorageReady Then
            Dim o_bytChunk() As Byte
            Dim o_lngItems As Long
            Dim o_lngPosition As Long
            Dim o_udtIndexRecord As typeIndexRecord
            Dim o_lngSize As Long
            
            o_udtIndexRecord = m_colIndex.Item(lngItemIndex)
            With o_udtIndexRecord
                Seek #m_intStorageFileNum, .lngStart
                
                If blnDecode Then
                    o_lngSize = GetOriginalSize(lngItemIndex)
                Else
                    o_lngSize = GetCompressedSize(lngItemIndex)
                End If
                
                ReDim bytContent(0 To o_lngSize - 1) As Byte
                
                m_ucmCRC.Clear
                
                o_lngPosition = 0
                
                For o_lngItems = LBound(.udtCompressedSection) To UBound(.udtCompressedSection)
                    ReDim o_bytChunk(0 To .udtCompressedSection(o_lngItems).lngCompressedSize - 1) As Byte
                    
                    Get #m_intStorageFileNum, , o_bytChunk
                    
                    m_ucmCRC.AddBytes o_bytChunk
                    
                    If blnDecode Then
                        Select Case .udeFlag
                            Case enumCompresssion.Zlib
                                m_ucmCompress.DecompressByteArray o_bytChunk, .udtCompressedSection(o_lngItems).lngOriginalSize
                            Case enumCompresssion.Huffman
                                m_ucmHuffman.DecodeByte o_bytChunk, UBound(o_bytChunk) + 1
                            Case enumCompresssion.Normal
                        End Select
                    End If

                    CopyMemory bytContent(o_lngPosition), o_bytChunk(0), UBound(o_bytChunk) + 1

                    o_lngPosition = o_lngPosition + UBound(o_bytChunk) + 1
                Next

                If Trim(m_ucmCRC.HexValue) = .strCRC Or Not blnDecode Then
                    o_blnResult = True
                Else
                    Err.Raise 101, "GetItemContent", "CRC failed!"
                    o_blnResult = False
                End If
            End With
        Else
            o_blnResult = False
        End If

        GetItemContent = o_blnResult
End Function


'get the actual content according to the index
'you have to do the decompression yourself if it is compressed.
Public Function ReadItemContentChunk(ByVal lngItemIndex As Long, ByRef bytContent() As Byte, ByVal lngOffset As Long, ByVal lngSize As Long, Optional ByVal blnDecode As Boolean = True) As Boolean
        Dim o_udeResult As enumReadChunkResult
        Dim o_udtIndexRecord As typeIndexRecord
        Dim o_lngItemSize As Long
        Dim o_lngItems As Long
        Dim o_lngPosition As Long
        Dim o_lngChunkSize As Long
        Dim o_lngChunckIndex As Long
        Dim o_lngChunkCount As Long
        Dim o_bytChunk() As Byte
        Dim o_blnResutl As Boolean
        
        If m_blnStorageReady Then
            o_udtIndexRecord = m_colIndex.Item(lngItemIndex)
            
            o_lngItemSize = GetOriginalSize(lngItemIndex)
            
            With o_udtIndexRecord
                If lngOffset >= 0 Then
                    o_blnResutl = lngOffset < o_lngItemSize
                Else
                    lngOffset = 0
                    If Loc(m_intStorageFileNum) < .lngStart Or Loc(m_intStorageFileNum) > .lngStart + o_lngItemSize - 1 Then
                        Seek #m_intStorageFileNum, .lngStart
                        o_blnResutl = True
                    Else
                        o_blnResutl = Loc(m_intStorageFileNum) - .lngStart <= o_lngItemSize
                    End If
                End If

                If o_blnResutl Then
                    If lngSize <= o_lngItemSize Then
                        If lngSize <= 0 Then lngSize = o_lngItemSize - lngOffset

                        o_lngChunckIndex = lngOffset \ m_lngChunkSize
                        If lngOffset / m_lngChunkSize >= 1 Then
                            o_lngChunckIndex = o_lngChunckIndex + 1
                        End If
                        
                        o_lngChunkCount = (lngOffset + lngSize) \ m_lngChunkSize
                        If (lngOffset + lngSize) / m_lngChunkSize >= 1 Then
                            o_lngChunkCount = o_lngChunkCount + 1
                        End If
                        If o_lngChunkCount = 0 Then o_lngChunkCount = 1
                        
                        If lngOffset > 0 Then
                            Seek #m_intStorageFileNum, .lngStart
                        End If

                        o_lngPosition = 0
                        
                        ReDim bytContent(0 To lngSize - 1) As Byte
                        
                        For o_lngItems = 0 To o_lngChunkCount - 1
                            ReDim o_bytChunk(0 To .udtCompressedSection(o_lngItems + o_lngChunckIndex).lngCompressedSize - 1) As Byte
                            
                            Get #m_intStorageFileNum, , o_bytChunk
                                                        
                            If blnDecode Then
                                Select Case .udeFlag
                                    Case enumCompresssion.Zlib
                                        m_ucmCompress.DecompressByteArray o_bytChunk, .udtCompressedSection(o_lngItems).lngOriginalSize
                                    Case enumCompresssion.Huffman
                                        m_ucmHuffman.DecodeByte o_bytChunk, UBound(o_bytChunk) + 1
                                    Case enumCompresssion.Normal
                                End Select
                            End If
                            
                            If o_lngItems = 0 Then
                                o_lngItemSize = m_lngChunkSize - (lngOffset Mod m_lngChunkSize)
                                CopyMemory bytContent(0), o_bytChunk(lngOffset Mod m_lngChunkSize + 1), o_lngItemSize
                                o_lngPosition = o_lngPosition + o_lngItemSize
                            ElseIf o_lngItems = o_lngChunkCount - 1 Then
                                o_lngItemSize = (lngSize - (m_lngChunkSize - (lngOffset Mod m_lngChunkSize) + 1)) Mod m_lngChunkSize
                                CopyMemory bytContent(o_lngPosition), o_bytChunk(0), o_lngItemSize
                                o_lngPosition = o_lngPosition + o_lngItemSize
                            Else
                                CopyMemory bytContent(o_lngPosition), o_bytChunk(0), UBound(o_bytChunk) + 1
                                o_lngPosition = o_lngPosition + UBound(o_bytChunk) + 1
                            End If
                        Next
                        
                        o_udeResult = enumReadChunkResult.OK
                    Else
                        o_udeResult = enumReadChunkResult.SizeToLarge
                    End If
                Else
                    o_udeResult = enumReadChunkResult.OffsetOutOfBound
                End If
            End With
        Else
            o_udeResult = enumReadChunkResult.StorageNotReady
        End If
        
        Erase o_bytChunk

        ReadItemContentChunk = o_udeResult
End Function


'get the actual plain text according to the index
Public Function GetItemText(ByVal lngItemIndex As Long) As String
        If m_blnStorageReady Then
            Dim o_bytContent() As Byte
            
            If GetItemContent(lngItemIndex, o_bytContent) Then
                GetItemText = StrConv(o_bytContent, vbUnicode)
            Else
                GetItemText = vbNullString
            End If
        Else
            GetItemText = vbNullString
        End If
End Function


'get the actual picture according to the index
Public Function GetItemPicture(ByVal lngItemIndex As Long) As IPicture
        If m_blnStorageReady Then
            Dim o_bytContent() As Byte
            
            If GetItemContent(lngItemIndex, o_bytContent) Then
                Set GetItemPicture = m_ucmDIBs.PictureFromByteStream(o_bytContent)
            Else
                Set GetItemPicture = Nothing
            End If
        Else
            Set GetItemPicture = Nothing
        End If
End Function


Private Sub RaiseError(ByVal lngCode As Long, ByVal strSource As String, ByVal strDescription As String)
        Err.Raise vbObjectError + lngCode, strSource, strDescription
End Sub


'save an item to the file
Public Function SaveItemToFile(ByVal lngItemIndex As Long, ByVal strFile As String) As Boolean
        Dim o_bytContent() As Byte
        
        If GetItemContent(lngItemIndex, o_bytContent) Then
            SaveItemToFile = m_ucmFile.SaveContentToFile(strFile, o_bytContent)
        Else
            SaveItemToFile = False
        End If
End Function


'since the actual content is still in there, remove them according to the index
Public Function VaccumStorage() As Boolean
        Dim o_blnResult As Boolean
        
        If m_blnStorageReady And m_blnUpdated Then
            Dim o_lngItems As Long
            Dim o_lngSections As Long
            Dim o_intFileNum As Integer
            Dim o_intFileNum2 As Integer
            Dim o_udtIndexRecord As typeIndexRecord
            Dim o_bytChunk() As Byte
            
            o_intFileNum = FreeFile
            Open m_strStorageFilePath & ".tmp" For Binary As #o_intFileNum
            
            o_intFileNum2 = FreeFile
            Open GetIndexFile & ".tmp" For Binary As #o_intFileNum2
            
            WriteHeader o_intFileNum2
            
            If m_colIndex.Count > 0 Then
                For o_lngItems = 1 To m_colIndex.Count
                    Put #o_intFileNum, Loc(o_intFileNum) + 1, m_bytSplitter
                    
                    o_udtIndexRecord = m_colIndex.Item(o_lngItems)
                                        
                    With o_udtIndexRecord
                        Seek #m_intStorageFileNum, .lngStart
                        
                        .lngStart = LOF(o_intFileNum) + 1
                        
                        For o_lngSections = LBound(.udtCompressedSection) To UBound(.udtCompressedSection)
                            ReDim o_bytChunk(0 To .udtCompressedSection(o_lngSections).lngCompressedSize - 1) As Byte
                            
                            Get #m_intStorageFileNum, , o_bytChunk

                            Put #o_intFileNum, Loc(o_intFileNum) + 1, o_bytChunk
                        Next

                    End With
                    
                    'o_udtIndexRecord.strCRC = Trim(m_ucmCRC.CalculateBytes(o_bytChunk))
                    Put #o_intFileNum2, Loc(o_intFileNum2) + 1, o_udtIndexRecord
                Next
            Else
            End If
            
            Close #o_intFileNum
            Close #o_intFileNum2
    
            CloseStorage
    
            FileCopy m_strStorageFilePath & ".tmp", m_strStorageFilePath
            Kill m_strStorageFilePath & ".tmp"
    
            FileCopy GetIndexFile() & ".tmp", GetIndexFile()
            Kill GetIndexFile() & ".tmp"
            
            o_blnResult = OpenStorage = enumOpenStorageResult.OK
        Else
            o_blnResult = False
        End If
        
        VaccumStorage = o_blnResult

End Function


'removes an item, simply removes it from the index for performance
Public Function DeleteItem(ByVal lngItemIndex As Long) As Boolean
        Dim o_blnResult As Boolean
        
        If m_blnStorageReady Then
            Dim o_intFileNum As Integer
            Dim o_lngItems As Long
            Dim o_udtIndexRecord As typeIndexRecord
            
            CloseStorage
            
            Kill GetIndexFile()
            
            CreateStorage
            
            o_intFileNum = FreeFile
            
            Open GetIndexFile() For Binary As #o_intFileNum
            
            Seek #o_intFileNum, LOF(o_intFileNum) + 1
            
            For o_lngItems = 1 To m_colIndex.Count
                If o_lngItems <> lngItemIndex Then
                    o_udtIndexRecord = m_colIndex.Item(o_lngItems)
                    Put #o_intFileNum, , o_udtIndexRecord
                End If
            Next
            
            Close #o_intFileNum
            
            o_blnResult = OpenStorage = enumOpenStorageResult.OK
        Else
            o_blnResult = False
        End If
        
        DeleteItem = o_blnResult
        
End Function


'add/update an file, any format!
Public Function AddUpdateItem(ByVal strKey As String, _
                              ByRef vntContent As Variant, _
                              ByVal strMemo As String, _
                              ByVal udeFlag As enumCompresssion, _
                              ByVal udeType As enumStorageType, _
                              ByVal lngItemData As Long, _
                              ByVal strCategory As String, _
                              ByVal blnAutoUpdate As Boolean) As enumAddUpdateItemResult
        Dim o_udeResult As enumAddUpdateItemResult
        
        m_ucmCRC.Clear

        If m_blnStorageReady Then
            If udeType = File Then
                If m_ucmFile.DoesFileExistEx(strKey) Then
                    o_udeResult = enumAddUpdateItemResult.OK
                Else
                    o_udeResult = enumAddUpdateItemResult.FileNotFound
                End If
            Else
                o_udeResult = enumAddUpdateItemResult.OK
            End If
            
            If o_udeResult = enumAddUpdateItemResult.OK Then
                Dim o_bytContent() As Byte
                Dim o_bytChunk() As Byte
                Dim o_lngItemIndex As Long
                Dim o_lngPosition As Long
                Dim o_lngSize As Long
                Dim o_initContentFileNum As Integer
                Dim o_lngChunkSize As Long
                Dim o_lngBound As Long
                Dim o_lngCompressedSize As Long
                Dim o_blnDone As Boolean
                Dim o_udtIndexRecord As typeIndexRecord
    
                If LOF(m_intStorageFileNum) > 0 Then
                    Seek #m_intStorageFileNum, LOF(m_intStorageFileNum) + 1
                End If
        
                Put #m_intStorageFileNum, , m_bytSplitter
    
                'find if the item already exists or not
                o_lngItemIndex = GetItemIndex(strKey)
                
                If o_lngItemIndex > 0 Then
                    'existing, update index
                    o_lngPosition = Loc(m_intStorageFileNum)
                    DeleteItem o_lngItemIndex
                    Seek #m_intStorageFileNum, o_lngPosition + 1
                Else
                    'Put #m_intIndexFileNum, , m_bytSplitter
                    Seek #m_intIndexFileNum, LOF(m_intIndexFileNum) + 1
                End If

                With o_udtIndexRecord
                    .lngStart = LOF(m_intStorageFileNum) + 1
                    
                    Select Case udeType
                        Case enumStorageType.File
                            o_lngSize = FileLen(strKey)
                            o_initContentFileNum = FreeFile
                            Open strKey For Binary As #o_initContentFileNum
                        Case enumStorageType.Text
                            o_bytContent = StrConv(vntContent, vbFromUnicode)
                        Case enumStorageType.ByteArray
                            o_bytContent = vntContent
                        Case enumStorageType.Image
                            m_ucmDIBs.CreateBitsFromStdPicture vntContent, o_bytContent
                    End Select
                    
                    If o_lngSize = 0 Then o_lngSize = UBound(o_bytContent) + 1
                    
                    o_lngChunkSize = m_lngChunkSize
                    
                    o_lngPosition = 0
                    
                    Do Until EOF(o_initContentFileNum) Or o_lngPosition >= o_lngSize Or o_blnDone
                        If o_lngPosition + o_lngChunkSize > o_lngSize Then
                            o_lngChunkSize = o_lngSize - o_lngPosition
                            o_blnDone = True
                        End If
                        
                        ReDim o_bytChunk(0 To o_lngChunkSize - 1)
                        
                        Select Case udeType
                            Case enumStorageType.File
                                Get #o_initContentFileNum, , o_bytChunk
                            Case enumStorageType.Text, enumStorageType.ByteArray, enumStorageType.Image
                                CopyMemory o_bytChunk(0), o_bytContent(o_lngPosition), o_lngChunkSize
                        End Select
                        
                        If o_lngPosition = 0 Then
                            o_lngBound = 0
                        Else
                            o_lngBound = UBound(.udtCompressedSection) + 1
                        End If
                        
                        ReDim Preserve .udtCompressedSection(o_lngBound) As typeCompressedSection
                        
                        .udtCompressedSection(o_lngBound).lngOriginalSize = UBound(o_bytChunk) + 1
                            
                        Select Case udeFlag
                            Case Zlib
                                m_ucmCompress.CompressByteArray o_bytChunk, COMPRESS_DEFAULT
                                o_lngCompressedSize = m_ucmCompress.CompressedSize
                            Case Huffman
                                m_ucmHuffman.EncodeByte o_bytChunk, UBound(o_bytChunk) + 1
                                o_lngCompressedSize = UBound(o_bytChunk) + 1
                            Case Normal
                                o_lngCompressedSize = .udtCompressedSection(o_lngBound).lngOriginalSize 'use original size since no compression
                        End Select
                        
                        m_ucmCRC.AddBytes o_bytChunk

                        Put #m_intStorageFileNum, , o_bytChunk
                                                
                        .udtCompressedSection(o_lngBound).lngCompressedSize = o_lngCompressedSize
                        
                        o_lngPosition = o_lngPosition + o_lngChunkSize
                    Loop
                    
                    Erase o_bytChunk
                    
                    Erase o_bytContent
                
                    .strKey = strKey
                                        
                    Select Case udeType
                        Case enumStorageType.File
                            .dteCreatedDate = FileDateTime(strKey)
                            .dteLastModifiedDate = FileDateTime(strKey) 'todo:
                            .strType = LCase(m_ucmFile.GetFileExtName(strKey))
                        Case enumStorageType.Text, enumStorageType.ByteArray, enumStorageType.Image
                            .dteCreatedDate = Now
                            .dteLastModifiedDate = Now
                    End Select
                    
                    .lngItemData = lngItemData
                    .strMemo = strMemo
                    .udeFlag = udeFlag
                    .strCRC = m_ucmCRC.HexValue
                End With
                
                Put #m_intIndexFileNum, , o_udtIndexRecord
                
                Close #o_initContentFileNum
                
                If blnAutoUpdate Then
                    'close & open again for content update
                    CloseStorage
    
                    If OpenStorage = enumOpenStorageResult.OK Then
                        o_udeResult = enumAddUpdateItemResult.OK
                    Else
                        o_udeResult = enumAddUpdateItemResult.Error
                    End If

                    m_blnUpdated = True
                Else
                    m_blnUpdated = False
                    
                    o_udeResult = True
                End If
            Else
            End If
        Else
            o_udeResult = False
        End If
        
        AddUpdateItem = o_udeResult
        
End Function


Private Sub Class_Initialize()
        m_bytIndexHeader(0) = 10
        m_bytIndexHeader(1) = 14
        m_bytIndexHeader(2) = 37
        m_bytIndexHeader(3) = 59
        m_bytIndexHeader(4) = 10
        m_bytIndexHeader(5) = 12
        
        m_bytSplitter(0) = 10
        m_bytSplitter(1) = 11
        m_bytSplitter(2) = 10
        m_bytSplitter(3) = 12
        m_bytSplitter(4) = 10
        m_bytSplitter(5) = 13
        
        m_lngChunkSize = mc_lngChunkSize
End Sub


Private Sub Class_Terminate()
        CloseStorage
        
        Set m_ucmFile = Nothing
        
        Set m_ucmCompress = Nothing
        
        Set m_ucmHuffman = Nothing
        
        Set m_ucmCRC = Nothing
End Sub

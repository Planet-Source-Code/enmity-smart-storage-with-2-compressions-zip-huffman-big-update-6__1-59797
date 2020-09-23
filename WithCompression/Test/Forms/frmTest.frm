VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Smart Storage Test"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   9990
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.TextBox txtCategory 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   8160
      TabIndex        =   35
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CheckBox chkFilters 
      Appearance      =   0  'Flat
      Caption         =   "Category"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   8280
      TabIndex        =   34
      Top             =   1680
      Width           =   1155
   End
   Begin VB.CheckBox chkFilters 
      Appearance      =   0  'Flat
      Caption         =   "Memo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   7560
      TabIndex        =   33
      Top             =   1680
      Width           =   795
   End
   Begin VB.CheckBox chkFilters 
      Appearance      =   0  'Flat
      Caption         =   "Flag"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   6840
      TabIndex        =   32
      Top             =   1680
      Width           =   795
   End
   Begin VB.CheckBox chkFilters 
      Appearance      =   0  'Flat
      Caption         =   "Type"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   6120
      TabIndex        =   31
      Top             =   1680
      Width           =   795
   End
   Begin VB.CheckBox chkFilters 
      Appearance      =   0  'Flat
      Caption         =   "ItemData"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   5040
      TabIndex        =   30
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CheckBox chkFilters 
      Appearance      =   0  'Flat
      Caption         =   "Updated"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   29
      Top             =   1680
      Width           =   975
   End
   Begin VB.CheckBox chkFilters 
      Appearance      =   0  'Flat
      Caption         =   "Created"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   28
      Top             =   1680
      Width           =   975
   End
   Begin VB.CheckBox chkFilters 
      Appearance      =   0  'Flat
      Caption         =   "Key"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   27
      Top             =   1680
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.TextBox txtFilter 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   840
      TabIndex        =   25
      ToolTipText     =   "Press Enter to filter!"
      Top             =   1560
      Width           =   1620
   End
   Begin VB.TextBox txtItemData 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5400
      TabIndex        =   23
      Top             =   1080
      Width           =   1500
   End
   Begin VB.ComboBox cboCompression 
      Height          =   300
      Left            =   7440
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtMemo 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   840
      TabIndex        =   20
      Top             =   1080
      Width           =   3660
   End
   Begin VB.CommandButton cmdCloseStorage 
      Caption         =   "C&lose"
      Height          =   375
      Left            =   6480
      TabIndex        =   19
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CheckBox chkAutoUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Auto"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8640
      TabIndex        =   18
      Top             =   600
      Value           =   1  'Checked
      Width           =   680
   End
   Begin prjTest.ucListView lvwRecords 
      Height          =   1695
      Left            =   840
      TabIndex        =   17
      ToolTipText     =   "double click item to save it!"
      Top             =   2040
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   2990
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   375
      Left            =   840
      TabIndex        =   16
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton cmdVaccum 
      Caption         =   "&Vaccum"
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9360
      TabIndex        =   14
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Del"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9360
      TabIndex        =   13
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox picContent 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   840
      ScaleHeight     =   2625
      ScaleWidth      =   8385
      TabIndex        =   11
      Top             =   3840
      Visible         =   0   'False
      Width           =   8415
   End
   Begin VB.TextBox txtContent 
      Height          =   2655
      Left            =   840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   3840
      Width           =   8415
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9360
      TabIndex        =   10
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtItemPath 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   600
      Width           =   6060
   End
   Begin VB.CommandButton cmdSelectItemFile 
      Caption         =   "..."
      Height          =   375
      Left            =   6900
      TabIndex        =   7
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelectDataFile 
      Caption         =   "..."
      Height          =   375
      Left            =   9360
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   7920
      TabIndex        =   4
      Top             =   6600
      Width           =   1335
   End
   Begin VB.TextBox txtDataPath 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   8460
   End
   Begin VB.CommandButton cmdOpenStorage 
      Caption         =   "&Open"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Categor&y:"
      Height          =   180
      Index           =   6
      Left            =   7200
      TabIndex        =   36
      Top             =   1080
      Width           =   810
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "&Filter:"
      Height          =   180
      Index           =   5
      Left            =   120
      TabIndex        =   26
      Top             =   1620
      Width           =   630
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Da&ta:"
      Height          =   180
      Index           =   4
      Left            =   4800
      TabIndex        =   24
      Top             =   1140
      Width           =   450
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "&Memo:"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   1140
      Width           =   450
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "&Item:"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   660
      Width           =   450
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "&Records:"
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "&DB:"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   270
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private m_ucmStorage As cStorage
Private m_ucmFile As New cFile
Private m_lngItems As Long



Private Sub cmdAdd_Click()
        With m_ucmStorage
            If .StorageReady Then
                If .AddUpdateItem(txtItemPath.Text, vbNull, txtMemo.Text, Choose(cboCompression.ListIndex + 1, enumCompresssion.Normal, enumCompresssion.Zlib, enumCompresssion.Huffman), enumStorageType.File, Val(txtItemData.Text), txtCategory.Text, chkAutoUpdate.Value = vbChecked) = enumOpenStorageResult.OK Then
                    OpenStorage
                    lvwRecords.ItemSelected(lvwRecords.Count - 1) = True
                    Me.Caption = "Storage Test (Items " & .Items.Count & ")"
                Else
                    MsgBox "failed to add/update item!"
                End If
            Else
                MsgBox "storage not ready!"
            End If
        End With
End Sub


Private Sub cmdCloseStorage_Click()
        CloseStorage
End Sub


Private Sub CloseStorage()
        m_ucmStorage.CloseStorage
        lvwRecords.Clear
        picContent.Picture = LoadPicture("")
        txtContent.Text = ""
        cmdAdd.Enabled = False
End Sub


Private Sub cmdCreate_Click()
        CloseStorage
        
        Set m_ucmStorage = New cStorage
                
        With m_ucmStorage
            .StorageFilePath = txtDataPath.Text
            MsgBox .CreateStorage()
            cmdAdd.Enabled = False
        End With
End Sub


Private Sub cmdDelete_Click()
        m_ucmStorage.DeleteItem GetItemIndex
        picContent.Cls
        txtContent.Text = ""
        OpenStorage
End Sub


Private Sub cmdExit_Click()
        Unload Me
End Sub


Private Sub OpenStorage()
        If Trim(txtDataPath.Text) <> "" Then
            CloseStorage
            
            Set m_ucmStorage = New cStorage
            
            With m_ucmStorage
                .StorageFilePath = txtDataPath.Text
                
                Select Case .OpenStorage
                    Case enumOpenStorageResult.OK
                        ShowItems .Items
                        
                        cmdAdd.Enabled = True
                    Case enumOpenStorageResult.VersionTooLow
                        MsgBox "version of storage file is too low!"
                    Case enumOpenStorageResult.VersionTooHigh
                        MsgBox "version of storage file is too hight!"
                    Case enumOpenStorageResult.Mailformed
                        MsgBox "wrong storage file!"
                    Case enumOpenStorageResult.Error
                        MsgBox "error in opening storage file!"
                    Case enumOpenStorageResult.StorageNotFound
                        MsgBox "storage file not found!"
                End Select
            End With
        Else
            MsgBox "Please select Storage file!"
        End If
End Sub


Private Sub ShowItems(ByRef colItems As Collection)
        Dim o_strType As String
        
        lvwRecords.Visible = False
        lvwRecords.Clear
        
        With colItems
            Me.Caption = "Smart Storage Test (Items " & .Count & ")"
            
            For m_lngItems = 1 To .Count
                With .Item(m_lngItems)
                    lvwRecords.ItemAdd m_lngItems, m_lngItems, 0, 0
                    Select Case .udeType
                        Case enumStorageType.File
                            o_strType = .strType
                        Case enumStorageType.Text
                            o_strType = "Text"
                        Case enumStorageType.ByteArray
                            o_strType = "ByteArray"
                        Case enumStorageType.Image
                            o_strType = "Image"
                    End Select
                    lvwRecords.SubItemSet m_lngItems - 1, 1, m_ucmFile.GetFileName(.strKey), 0
                    lvwRecords.SubItemSet m_lngItems - 1, 2, o_strType, 0
                    lvwRecords.SubItemSet m_lngItems - 1, 3, m_ucmStorage.GetOriginalSize(m_lngItems), 0
                    lvwRecords.SubItemSet m_lngItems - 1, 4, m_ucmStorage.GetCompressedSize(m_lngItems), 0
                    lvwRecords.SubItemSet m_lngItems - 1, 5, .strCRC, 0
                    lvwRecords.SubItemSet m_lngItems - 1, 6, m_ucmStorage.GetCompressionTypeName(.udeFlag), 0
                    lvwRecords.SubItemSet m_lngItems - 1, 7, m_ucmStorage.GetStorageTypeName(.udeType), 0
                    lvwRecords.SubItemSet m_lngItems - 1, 8, .strCategory, 0
                    lvwRecords.SubItemSet m_lngItems - 1, 9, .strKey, 0
                    lvwRecords.SubItemSet m_lngItems - 1, 10, .strMemo, 0
                End With
            Next
        End With
        
        lvwRecords.Visible = True
End Sub


Private Sub cmdOpenStorage_Click()
        OpenStorage
End Sub


Private Sub cmdSave_Click()
        SaveItem
End Sub


Private Sub cmdSelectDataFile_Click()
        SelectFile "Storage File", "*.sdb", txtDataPath
End Sub


Private Sub cmdSelectItemFile_Click()
        SelectFile "Item File", "*.*", txtItemPath
        If txtItemPath.Text <> "" Then
            ShowItem True
        End If
End Sub


Private Sub SelectFile(ByVal strFilterName As String, ByVal strFilterExt As String, ByRef txtItem As TextBox)
        With New cFileDlg
            .Filters.AddFilter strFilterName, strFilterExt
            .flags = ofnPathMustExist Or ofnFileMustExist
            '.InitDir = m_ucmFile.GetFilePath(txtDataPath.Text)
            If .ShowOpen(Me.hWnd) Then
                txtItem.Text = .FileName
            End If
        End With
End Sub


Private Sub cmdTest_Click()
'        With New cCompression
'            .CompressFile App.Path & "\..\..\TestData\Animales peligrosos-Papel tapiz.jpg", App.Path & "\..\..\TestData\Animales peligrosos-Papel tapiz.jpg.comp", COMPRESS_DEFAULT
'            .DecompressFile App.Path & "\..\..\TestData\Animales peligrosos-Papel tapiz.jpg.comp", App.Path & "\..\..\TestData\Animales peligrosos-Papel tapiz2.jpg"
'        End With
'        With New cHuffman
'            .EncodeFile App.Path & "\..\..\TestData\Animales peligrosos-Papel tapiz.jpg", App.Path & "\..\..\TestData\Animales peligrosos-Papel tapiz.jpg.comp"
'            .DecodeFile App.Path & "\..\..\TestData\Animales peligrosos-Papel tapiz.jpg.comp", App.Path & "\..\..\TestData\Animales peligrosos-Papel tapiz2.jpg"
'        End With
        Dim o_bytContent() As Byte
        
        With m_ucmStorage
            Debug.Print .ReadItemContentChunk(GetItemIndex(), o_bytContent, 1000, 0)
            MsgBox StrConv(o_bytContent, vbUnicode)
        End With
End Sub


Private Sub cmdVaccum_Click()
        MsgBox m_ucmStorage.VaccumStorage
End Sub


Private Sub Form_Load()
        If Not m_ucmFile.DoesFileExistEx(App.Path & "\zlib.dll") Then
            Dim o_bytData() As Byte
            o_bytData = LoadResData("Data", "CUSTOM")
            m_ucmFile.SaveContentToFile App.Path & "\zlib.dll", o_bytData
            Erase o_bytData
        End If
        
        txtDataPath.Text = App.Path & "\Test.sdb"
        txtItemPath.Text = App.Path & "\..\..\TestData\Animales peligrosos-Papel tapiz.jpg"
        
        With cboCompression
            .AddItem "Normal"
            .AddItem "Zip"
            .AddItem "Huffman"
            .ListIndex = 1
        End With
        
        Set m_ucmStorage = New cStorage
            
        With lvwRecords
        
            Call .Initialize
            
            Call .InitializeImageListSmall
            Call .InitializeImageListLarge
            Call .InitializeImageListHeader
            Call .ImageListSmall_AddBitmap(LoadResPicture("IL16x16", vbResBitmap), vbMagenta)
            Call .ImageListLarge_AddBitmap(LoadResPicture("IL32x32", vbResBitmap), vbMagenta)
            Call .ImageListHeader_AddBitmap(LoadResPicture("ILHEADER", vbResBitmap), vbMagenta)
            
            Call .ColumnAdd(0, "Index", 40, [caLeft])
            Call .ColumnAdd(1, "File Name", 80, [caLeft])
            Call .ColumnAdd(2, "Ext", 40, [caLeft])
            Call .ColumnAdd(3, "Orginal", 80, [caRight])
            Call .ColumnAdd(4, "Compressed", 80, [caRight])
            Call .ColumnAdd(5, "CRC", 80, [caLeft])
            Call .ColumnAdd(6, "Flag", 40, [caLeft])
            Call .ColumnAdd(7, "Type", 40, [caLeft])
            Call .ColumnAdd(8, "Category", 80, [caRight])
            Call .ColumnAdd(9, "Path", 200, [caLeft])
            Call .ColumnAdd(10, "Memo", 100, [caLeft])

            .ViewMode = vmDetails
            .BorderStyle = bsThin
            .FullRowSelect = True
            .GridLines = True
            .HeaderDragDrop = True
            .HeaderFlat = True
            .HideSelection = False
            .OneClickActivate = True
            .ScrollBarFlat = True
            .TrackSelect = True
            .UnderlineHot = True
            .Visible = True
        End With
            
End Sub


Private Sub Form_Unload(Cancel As Integer)
        Set m_ucmStorage = Nothing
        Set m_ucmFile = Nothing
End Sub


Private Sub lvwRecords_ItemClick(Item As Integer)
        ShowItem False
End Sub


Private Sub lvwRecords_DBLClick()
        SaveItem
End Sub


Private Function GetItemIndex() As Long
        GetItemIndex = Val(lvwRecords.SubItemText(lvwRecords.SelectedItem, 0))
End Function


Private Sub ShowItem(ByVal blnDirect As Boolean)
    On Error GoTo HandleError
    
        With m_ucmStorage
            Dim o_blnResult As Boolean
            
            If blnDirect Then
                o_blnResult = blnDirect
            Else
                o_blnResult = .Items.Count > 0
            End If
            
            If o_blnResult Then
                cmdDelete.Enabled = True
                cmdSave.Enabled = True
                
                Dim o_strType As String
                
                If blnDirect Then
                    txtMemo.Text = txtItemPath.Text
                    o_strType = m_ucmFile.GetFileExtName(txtItemPath.Text)
                    txtItemData.Text = ""
                Else
                    o_strType = .Items.Item(GetItemIndex).strType
                    txtMemo.Text = .Items.Item(GetItemIndex).strMemo
                    txtItemData.Text = .Items.Item(GetItemIndex).lngItemData
                End If
                
                Select Case LCase(o_strType)
                    Case "jpeg", "jpg", "bmp", "gif", "wmf", "ico", "rle"
                        If blnDirect Then
                            Set picContent.Picture = LoadPicture(txtItemPath.Text)
                        Else
                            Set picContent.Picture = .GetItemPicture(GetItemIndex)
                        End If
                        txtContent.Visible = False
                        picContent.Visible = True
                    Case "txt", "pif", "bat", "htm", "html", "js", "dhtml", "asp", "jsp", "vbp", "frm", "cls", "bas", "vb", "cs", "c", "h", "cpp", "hpp"
                        txtContent.Text = ""
                        If blnDirect Then
                            txtContent.Text = m_ucmFile.LoadTextFromFile(txtItemPath.Text)
                        Else
                            txtContent.Text = .GetItemText(GetItemIndex)
                        End If
                        picContent.Visible = False
                        txtContent.Visible = True
                End Select
            Else
                cmdDelete.Enabled = False
                cmdSave.Enabled = False
            End If
        End With
    
    Exit Sub
    
HandleError:
    MsgBox Err.Source & ":" & Err.Description

End Sub


Private Sub SaveItem()
        If m_ucmStorage.Items.Count > 0 And lvwRecords.SelectedItem > -1 Then
            With New cFileDlg
                .Filters.AddFilter m_ucmStorage.Items(GetItemIndex).strType & "ÎÄ¼þ", "*." & m_ucmStorage.Items(GetItemIndex).strType
                .FilterIndex = 2
                .flags = ofnPathMustExist Or ofnOverwritePrompt
                .FileName = m_ucmFile.GetFileName(m_ucmStorage.Items(GetItemIndex).strKey)
                If .ShowSave(Me.hWnd) Then
                    If InStrRev(.FileName, "." & m_ucmStorage.Items(GetItemIndex).strType) = 0 Then
                        .FileName = .FileName & "." & m_ucmStorage.Items(GetItemIndex).strType
                    End If
                    
                    m_ucmStorage.SaveItemToFile GetItemIndex, .FileName

                End If
            End With
        End If
End Sub


Private Sub txtFilter_KeyPress(KeyAscii As Integer)
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
            Dim o_udeFilter As enumFilterType
            If chkFilters(0).Value = vbChecked Then o_udeFilter = o_udeFilter Or Key
            If chkFilters(1).Value = vbChecked Then o_udeFilter = o_udeFilter Or CreatedDate
            If chkFilters(2).Value = vbChecked Then o_udeFilter = o_udeFilter Or LastModifiedDate
            If chkFilters(3).Value = vbChecked Then o_udeFilter = o_udeFilter Or ItemData
            If chkFilters(4).Value = vbChecked Then o_udeFilter = o_udeFilter Or [Type]
            If chkFilters(5).Value = vbChecked Then o_udeFilter = o_udeFilter Or Flag
            If chkFilters(6).Value = vbChecked Then o_udeFilter = o_udeFilter Or Memo
            If chkFilters(7).Value = vbChecked Then o_udeFilter = o_udeFilter Or Category
            ShowItems m_ucmStorage.FilterItems(txtFilter.Text, o_udeFilter)
        End If
End Sub

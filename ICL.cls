VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ===========================================================================
' Filename:    ICL.cls
' Author:      Rob Anderson (randers99@hotmail.com)
' Date:        May 15, 2003
'
' Requires:    None
'
' Description:
' A class which reads and writes Icon Libraries, as well as imports and exports
' icons files from those libraries.
'
' An ICL file is basically a 16-bit (NE) DLL with just Icon Resource information
' within it. Any thoughts, critisms, and/or bugs can be directed at
' randers99@hotmail.com.
'
' ===========================================================================

Private Const mcstrMODULE_BASE_NAME As String = ".ICL."

' Property variables
Private mstrFileName As String

' Internal Module level variables
Private maudtIconData() As IconData

Private mlngIconCount   As Long

' API Declarations
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As FileOffsetTypes) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Private Declare Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As Any, ByVal wUsage As Long) As Long

Private Const RT_ICON       As Long = &H3
Private Const RT_GROUP_ICON As Long = &HE

Private Const RT_HIGHORDER  As Long = &H8000&

Private Const GENERIC_READ              As Long = &H80000000
Private Const GENERIC_WRITE             As Long = &H40000000
Private Const FILE_SHARE_READ           As Long = &H1
Private Const FILE_SHARE_WRITE          As Long = &H2
Private Const CREATE_ALWAYS             As Long = 2
Private Const CREATE_NEW                As Long = 1
Private Const OPEN_ALWAYS               As Long = 4
Private Const OPEN_EXISTING             As Long = 3
Private Const TRUNCATE_EXISTING         As Long = 5
Private Const FILE_ATTRIBUTE_ARCHIVE    As Long = &H20
Private Const FILE_ATTRIBUTE_HIDDEN     As Long = &H2
Private Const FILE_ATTRIBUTE_NORMAL     As Long = &H80
Private Const FILE_ATTRIBUTE_READONLY   As Long = &H1
Private Const FILE_ATTRIBUTE_SYSTEM     As Long = &H4
Private Const FILE_FLAG_DELETE_ON_CLOSE As Long = &H4000000
Private Const FILE_FLAG_NO_BUFFERING    As Long = &H20000000
Private Const FILE_FLAG_OVERLAPPED      As Long = &H40000000
Private Const FILE_FLAG_POSIX_SEMANTICS As Long = &H1000000
Private Const FILE_FLAG_RANDOM_ACCESS   As Long = &H10000000
Private Const FILE_FLAG_SEQUENTIAL_SCAN As Long = &H8000000
Private Const FILE_FLAG_WRITE_THROUGH   As Long = &H80000000

' Colour table information
Private Const DIB_PAL_COLORS        As Long = 1 '  color table in palette indices
Private Const DIB_PAL_INDICES       As Long = 2 '  No color table indices into surf palette
Private Const DIB_PAL_LOGINDICES    As Long = 4 '  No color table indices into DC palette
Private Const DIB_PAL_PHYSINDICES   As Long = 2 '  No color table indices into surf palette
Private Const DIB_RGB_COLORS        As Long = 0 '  color table in RGBs

' Resource idTypes
Private Const IMAGE_BITMAP  As Long = 0
Private Const IMAGE_ICON    As Long = 1
Private Const IMAGE_CURSOR  As Long = 2

' Bitmap compression types
Private Const BI_RGB        As Long = 0&
Private Const BI_RLE4       As Long = 2&
Private Const BI_RLE8       As Long = 1&

Private Const BITMAP_HEADER_SIZE    As Long = 40
Private Const MEMICONDIRENTRY_SIZE  As Long = 14

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hBmMask As Long
    hbmColor As Long
End Type

Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type IMAGE_DOS_HEADER
    e_magic As Integer      ' Magic number
    e_cblp As Integer       ' Bytes on last page of file
    e_cp As Integer         ' Pages in file
    e_crlc As Integer       ' Relocations
    e_cparhdr As Integer    ' Size of header in paragraphs
    e_minalloc As Integer   ' Minimum extra paragraphs needed
    e_maxalloc As Integer   ' Maximum extra paragraphs needed
    e_ss As Integer         ' Initial (relative) SS value
    e_sp As Integer         ' Initial SP value
    e_csum As Integer       ' Checksum
    e_ip As Integer         ' Initial IP value
    e_cs As Integer         ' Initial (relative) CS value
    e_lfarlc As Integer     ' File address of relocation table
    e_ovno As Integer       ' Overlay number
    e_res(0 To 3) As Integer ' Reserved words
    e_oemid As Integer      ' OEM identifier (for e_oeminfo)
    e_oeminfo As Integer    ' OEM information; e_oemid specific
    e_res2(0 To 9) As Integer ' Reserved words
    e_lfanew As Long        ' File address of new exe header
End Type

Private Type IMAGE_WINDOWS_NE_HEADER
    Machine As Integer ' "NE"
    LinkerVersion As Byte
    LinkerRevision As Byte
    EntryOffset As Integer
    EntryLength As Integer
    Reserved1 As Long
    Flags As Integer
    SegNumber As Integer
    InitialHeapSize As Integer
    InitialStackSize As Integer
    SegOffsetCSIP As Long
    SegOffsetSSSP As Long
    SegmentTableCount As Integer
    ModRefTableCount As Integer
    NonResidentNameSize As Integer
    SegmentTableOffset As Integer
    ResourceTableOffset As Integer
    ResidentNameTableOffset As Integer
    ModuleRefTableOffset As Integer
    ImportedNameTableOffset As Integer
    NonResidentNameTableOffset As Long
    MovableEntryPointCount As Integer
    ShiftCount As Integer
    ResourceSegmentCount As Integer
    TargetOS As Byte
    OSAddlInfo As Byte
    FastLoadOffset As Integer
    FastLoadLength As Integer
    Reserved2 As Integer
    WindowsVersion As Integer
End Type

Private Type TNAMEINFO
    rnOffset As Integer
    rnLength As Integer
    rnFlags As Integer
    rnID As Integer
    rnHandle As Integer
    rnUsage As Integer
End Type

Private Enum FileOffsetTypes
    FILE_BEGIN = 0&
    FILE_CURRENT = 1&
    FILE_END = 2&
End Enum

Private Type ICONDIRENTRY
    Width As Byte               ' Width of the image
    Height As Byte              ' Height of the image (times 2)
    ColorCount As Byte          ' Number of colors in image (0 if >=8bpp)
    Reserved As Byte            ' Reserved
    Planes As Integer           ' Color Planes
    BitsPerPixel As Integer     ' Bits per pixel
    DataSize As Long            ' Which icon corresponds to this entry?
    DataOffset As Long          ' where in the file is this image
End Type

Private Type IconEntry
    Width As Byte
    Height As Byte
    NumColors As Byte
    Reserved As Byte
    NumPlanes As Integer
    BitsPerPixel As Integer
    DataSize As Long
    DataOffset As Long
End Type

Private Type ICONDIR
    idReserved As Integer       ' Reserved
    idType As Integer           ' resource type (1 for icons)
    idCount As Integer          ' how many images?
End Type

Private Type BITMAPINFOHEADER   ' 40 bytes
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type MEMICONDIRENTRY
    bWidth As Byte               ' Width of the image
    bHeight As Byte              ' Height of the image (times 2)
    bColorCount As Byte          ' Number of colors in image (0 if >=8bpp)
    bReserved As Byte            ' Reserved
    wPlanes As Integer           ' Color Planes
    wBitCount As Integer         ' Bits per pixel
    dwBytesInRes As Long         ' how many bytes in this resource?
    nID As Integer               ' the ID
End Type

' Custom UDTs
Private Type IconGroupData
    DirHeader As ICONDIR
    Entries() As MEMICONDIRENTRY
End Type

Private Type IconBitmapData
    Header As BITMAPINFOHEADER
    Palette() As RGBQUAD
    XORMap() As Byte
    ANDMap() As Byte
End Type

Private Type IconData
    Name As String
    Group As IconGroupData
    Bitmaps() As IconBitmapData
End Type

Private Type mudtErrorData
    AddlInfo As String
    Desc As String
    HelpContext As String
    HelpFile As String
    LineNumber As Long
    Method As String
    Number As Long
    Source As String
End Type

Private Type DataOffsetLength
    offset As Long
    Length As Long
End Type

' The size of the BITMAPINFO structure depends on whether there is a colour
' table in the DIB or not and its size.  For mono (1bpp) DIBs, the colour
' table has two entries, for 16 colour (4bpp) DIBs, there are 16 entries,
' for 256 colour (8bpp) DIBs there are 256 entries, otherwise the DIB is
' 24 bits per pixel and has no colour table.

Private Type BITMAPINFO_1BPP
   bmiHeader As BITMAPINFOHEADER
   bmiColors(0 To 1) As RGBQUAD
End Type
Private Type BITMAPINFO_8BPP
   bmiHeader As BITMAPINFOHEADER
   bmiColors(0 To 255) As RGBQUAD
End Type

' <-------------------------------------------------------------------------------------------------->

Private Enum menmErrorNumber
    No_Error = 0
    ERR_INVALID_PROPERTY_ASSIGN = 1000
    ERR_INVALID_FILENAME = 2000
    ERR_WRITING_FILE = 2500
    ERR_READING_FILE = 2500
    ERR_INVALID_DATA_FORMAT = 3000
    ERR_DATA_OUT_OF_RANGE = 3001
    ERR_INVALID_PARAMETER = 3002
    ERR_FILEIO = 5000
End Enum

Private Enum OpenMode
    ForReading = 0
    ForWriting = 1
End Enum

Private Enum ECFIImageConstants
   ecfiImage = 0
   ecfiMask = 1
End Enum

' <-------------------------------------------------------------------------------------------------->

Public Property Get Filename() As String
    Filename = mstrFileName
End Property

Public Property Let Filename(ByVal Value As String)
    mstrFileName = Value
End Property

' <-------------------------------------------------------------------------------------------------->
'
'     Module Name:  ICL.IconCount (Public)
'
'     Description:  Total number of Icons found in the Library. This is one larger than the maximum index of maudtIconData().
'
'          Inputs:  (None)
'
'         Outputs:  Long
'
'  Change History:
'
'           Initials+#  Reference   Tracking #  Date        Description
'           ----------  ---------   ----------  ----------  ----------------------------------------------------
'           rda001      ICL         n/a         05/27/2003  Created for Project
'
Public Property Get IconCount() As Long
    IconCount = mlngIconCount
End Property

' <-------------------------------------------------------------------------------------------------->
'
'     Module Name:  ICL.ImageCount (Public)
'
'     Description:  Total number of images found in a given Icon.
'
'          Inputs:  Index, ByVal Long
'
'         Outputs:  Long
'
'  Change History:
'
'           Initials+#  Reference   Tracking #  Date        Description
'           ----------  ---------   ----------  ----------  ----------------------------------------------------
'           rda001      ICL         n/a         05/27/2003  Created for Project
'
Public Property Get ImageCount(ByVal Index As Long) As Long

    On Error GoTo ErrorCode

Dim udtErrInfo  As mudtErrorData

    udtErrInfo.Method = App.Title & mcstrMODULE_BASE_NAME & "Property Get ImageCount()"
    
    ' Make sure the Index passed in is valid.
    If ((Index < 1) Or (Index > mlngIconCount)) Then
        Err.Raise menmErrorNumber.ERR_DATA_OUT_OF_RANGE, udtErrInfo.Method, "Invalid Icon Index."
    End If
    
    ImageCount = maudtIconData(Index - 1).Group.DirHeader.idCount
    
' <---------------------------------------------------------->
ExitCode:

    ' Deallocate all object references.

    ' If an error occurs, raise the Err object and exit the function.
    If (udtErrInfo.Number <> menmErrorNumber.No_Error) Then
        Err.Raise udtErrInfo.Number, udtErrInfo.Source, udtErrInfo.Desc
    End If

    Exit Property

' <---------------------------------------------------------->
ErrorCode:

    ' Set a permanent break point for the IDE. This has no effect in compiled code.
    Debug.Assert False

    ' Error occurred store it, stop error catching, and go to exit
    Call SetErrorInformation(Err.Number, Err.Source, Err.Description, Erl(), udtErrInfo)
    On Error GoTo 0
    Resume ExitCode
    Resume ' For Debug Only!

End Property

' <-------------------------------------------------------------------------------------------------->
'
'     Module Name:  ICL.ImageWidth (Public)
'
'     Description:  Width of a given image.
'
'          Inputs:  IconIndex,  ByVal Long
'                   ImageIndex, ByVal Long
'
'         Outputs:  Long
'
'  Change History:
'
'           Initials+#  Reference   Tracking #  Date        Description
'           ----------  ---------   ----------  ----------  ----------------------------------------------------
'           rda001      ICL         n/a         05/28/2003  Created for Project
'
Public Property Get ImageWidth(ByVal IconIndex As Long, ByVal ImageIndex As Long) As Long

    On Error GoTo ErrorCode

Dim udtErrInfo  As mudtErrorData

    udtErrInfo.Method = App.Title & mcstrMODULE_BASE_NAME & "Property Get ImageWidth()"
    
    ' Make sure the Icon Index passed in is valid.
    If ((IconIndex < 1) Or (IconIndex > mlngIconCount)) Then
        Err.Raise menmErrorNumber.ERR_DATA_OUT_OF_RANGE, udtErrInfo.Method, "Invalid Icon Index."
    End If
    
    ' Make sure the Image Index passed in is valid.
    If ((ImageIndex < 1) Or (ImageIndex > ImageCount(IconIndex))) Then
        Err.Raise menmErrorNumber.ERR_DATA_OUT_OF_RANGE, udtErrInfo.Method, "Invalid Image Index."
    End If
    
    ImageWidth = maudtIconData(IconIndex - 1).Bitmaps(ImageIndex - 1).Header.biWidth
    
' <---------------------------------------------------------->
ExitCode:

    ' Deallocate all object references.

    ' If an error occurs, raise the Err object and exit the function.
    If (udtErrInfo.Number <> menmErrorNumber.No_Error) Then
        Err.Raise udtErrInfo.Number, udtErrInfo.Source, udtErrInfo.Desc
    End If

    Exit Property

' <---------------------------------------------------------->
ErrorCode:

    ' Set a permanent break point for the IDE. This has no effect in compiled code.
    Debug.Assert False

    ' Error occurred store it, stop error catching, and go to exit
    Call SetErrorInformation(Err.Number, Err.Source, Err.Description, Erl(), udtErrInfo)
    On Error GoTo 0
    Resume ExitCode
    Resume ' For Debug Only!

End Property

' <-------------------------------------------------------------------------------------------------->
'
'     Module Name:  ICL.ImageHeight (Public)
'
'     Description:  Height of a given image.
'
'          Inputs:  IconIndex,  ByVal Long
'                   ImageIndex, ByVal Long
'
'         Outputs:  Long
'
'  Change History:
'
'           Initials+#  Reference   Tracking #  Date        Description
'           ----------  ---------   ----------  ----------  ----------------------------------------------------
'           rda001      ICL         n/a         05/28/2003  Created for Project
'
Public Property Get ImageHeight(ByVal IconIndex As Long, ByVal ImageIndex As Long) As Long

    On Error GoTo ErrorCode

Dim udtErrInfo  As mudtErrorData

    udtErrInfo.Method = App.Title & mcstrMODULE_BASE_NAME & "Property Get ImageHeight()"
    
    ' Make sure the Icon Index passed in is valid.
    If ((IconIndex < 1) Or (IconIndex > mlngIconCount)) Then
        Err.Raise menmErrorNumber.ERR_DATA_OUT_OF_RANGE, udtErrInfo.Method, "Invalid Icon Index."
    End If
    
    ' Make sure the Image Index passed in is valid.
    If ((ImageIndex < 1) Or (ImageIndex > ImageCount(IconIndex))) Then
        Err.Raise menmErrorNumber.ERR_DATA_OUT_OF_RANGE, udtErrInfo.Method, "Invalid Image Index."
    End If
    
    ImageHeight = maudtIconData(IconIndex - 1).Bitmaps(ImageIndex - 1).Header.biHeight \ 2
    
' <---------------------------------------------------------->
ExitCode:

    ' Deallocate all object references.

    ' If an error occurs, raise the Err object and exit the function.
    If (udtErrInfo.Number <> menmErrorNumber.No_Error) Then
        Err.Raise udtErrInfo.Number, udtErrInfo.Source, udtErrInfo.Desc
    End If

    Exit Property

' <---------------------------------------------------------->
ErrorCode:

    ' Set a permanent break point for the IDE. This has no effect in compiled code.
    Debug.Assert False

    ' Error occurred store it, stop error catching, and go to exit
    Call SetErrorInformation(Err.Number, Err.Source, Err.Description, Erl(), udtErrInfo)
    On Error GoTo 0
    Resume ExitCode
    Resume ' For Debug Only!

End Property

' <-------------------------------------------------------------------------------------------------->
'
'     Module Name:  ICL.ColorDepth (Public)
'
'     Description:  Color Depth of a given image.
'
'          Inputs:  IconIndex,  ByVal Long
'                   ImageIndex, ByVal Long
'
'         Outputs:  Long
'
'  Change History:
'
'           Initials+#  Reference   Tracking #  Date        Description
'           ----------  ---------   ----------  ----------  ----------------------------------------------------
'           rda001      ICL         n/a         05/28/2003  Created for Project
'
Public Property Get ColorDepth(ByVal IconIndex As Long, ByVal ImageIndex As Long) As Long

    On Error GoTo ErrorCode

Dim udtErrInfo  As mudtErrorData

    udtErrInfo.Method = App.Title & mcstrMODULE_BASE_NAME & "Property Get ColorDepth()"

    ' Make sure the Icon Index passed in is valid.
    If ((IconIndex < 1) Or (IconIndex > mlngIconCount)) Then
        Err.Raise menmErrorNumber.ERR_DATA_OUT_OF_RANGE, udtErrInfo.Method, "Invalid Icon Index."
    End If
    
    ' Make sure the Image Index passed in is valid.
    If ((ImageIndex < 1) Or (ImageIndex > ImageCount(IconIndex))) Then
        Err.Raise menmErrorNumber.ERR_DATA_OUT_OF_RANGE, udtErrInfo.Method, "Invalid Image Index."
    End If
    
    ColorDepth = maudtIconData(IconIndex - 1).Bitmaps(ImageIndex - 1).Header.biBitCount
    
' <---------------------------------------------------------->
ExitCode:

    ' Deallocate all object references.

    ' If an error occurs, raise the Err object and exit the function.
    If (udtErrInfo.Number <> menmErrorNumber.No_Error) Then
        Err.Raise udtErrInfo.Number, udtErrInfo.Source, udtErrInfo.Desc
    End If

    Exit Property

' <---------------------------------------------------------->
ErrorCode:

    ' Set a permanent break point for the IDE. This has no effect in compiled code.
    Debug.Assert False

    ' Error occurred store it, stop error catching, and go to exit
    Call SetErrorInformation(Err.Number, Err.Source, Err.Description, Erl(), udtErrInfo)
    On Error GoTo 0
    Resume ExitCode
    Resume ' For Debug Only!

End Property

' <-------------------------------------------------------------------------------------------------->
'
'     Module Name:  ICL.IconName (Public)
'
'     Description:  Name of the given Icon
'
'          Inputs:  IconIndex,  ByVal Long
'
'         Outputs:  Long
'
'  Change History:
'
'           Initials+#  Reference   Tracking #  Date        Description
'           ----------  ---------   ----------  ----------  ----------------------------------------------------
'           rda001      ICL         n/a         06/02/2003  Created for Project
'
Public Property Get IconName(ByVal IconIndex As Long) As String

    On Error GoTo ErrorCode

Dim udtErrInfo  As mudtErrorData

    udtErrInfo.Method = App.Title & mcstrMODULE_BASE_NAME & "Property Get IconName()"

    ' Make sure the Icon Index passed in is valid.
    If ((IconIndex < 1) Or (IconIndex > mlngIconCount)) Then
        Err.Raise menmErrorNumber.ERR_DATA_OUT_OF_RANGE, udtErrInfo.Method, "Invalid Icon Index."
    End If
    
    IconName = maudtIconData(IconIndex - 1).Name
    
' <---------------------------------------------------------->
ExitCode:

    ' Deallocate all object references.

    ' If an error occurs, raise the Err object and exit the function.
    If (udtErrInfo.Number <> menmErrorNumber.No_Error) Then
        Err.Raise udtErrInfo.Number, udtErrInfo.Source, udtErrInfo.Desc
    End If

    Exit Property

' <---------------------------------------------------------->
ErrorCode:

    ' Set a permanent break point for the IDE. This has no effect in compiled code.
    Debug.Assert False

    ' Error occurred store it, stop error catching, and go to exit
    Call SetErrorInformation(Err.Number, Err.Source, Err.Description, Erl(), udtErrInfo)
    On Error GoTo 0
    Resume ExitCode
    Resume ' For Debug Only!

End Property

' <-------------------------------------------------------------------------------------------------->
'
'     Module Name:  ICL.OpenFile (Private)
'
'     Description:  Opens the file from the module level filename variable.
'
'          Inputs:  vstrFileName, ByVal String
'                   vlngMode,     ByVal OpenMode [Long]
'
'         Outputs:  Returns the API File handle.
'
'  Change History:
'
'           Initials+#  Reference   Tracking #  Date        Description
'           ----------  ---------   ----------  ----------  ----------------------------------------------------
'           rda001      ICL         n/a         05/15/2003  Created for Project
'
Private Function OpenFile(ByVal vstrFilename As String, ByVal vlngMode As OpenMode) As Long

    On Error GoTo ErrorCode

Dim udtErrInfo  As mudtErrorData
Dim lngHandle   As Long

    udtErrInfo.Method = App.Title & mcstrMODULE_BASE_NAME & "OpenFile()"

    If (vlngMode = OpenMode.ForReading) Then
        lngHandle = CreateFile(vstrFilename, GENERIC_READ, FILE_SHARE_READ, ByVal CLng(0), _
                               OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
    Else
        lngHandle = CreateFile(vstrFilename, GENERIC_WRITE, 0, 0, _
                            CREATE_ALWAYS, FILE_ATTRIBUTE_ARCHIVE, 0)
    End If
    
    If (lngHandle = -1) Then
        Err.Raise menmErrorNumber.ERR_READING_FILE, udtErrInfo.Method, "Could not open the file successfully."
    End If

    OpenFile = lngHandle

' <---------------------------------------------------------->
ExitCode:

    ' Deallocate all object references.

    ' If an error occurs, raise the Err object and exit the function.
    If (udtErrInfo.Number <> menmErrorNumber.No_Error) Then
        Err.Raise udtErrInfo.Number, udtErrInfo.Source, udtErrInfo.Desc
    End If

    Exit Function

' <---------------------------------------------------------->
ErrorCode:

    ' Set a permanent break point for the IDE. This has no effect in compiled code.
    Debug.Assert False

    ' Error occurred store it, stop error catching, and go to exit
    Call SetErrorInformation(Err.Number, Err.Source, Err.Description, Erl(), udtErrInfo)
    On Error GoTo 0
    Resume ExitCode
    Resume ' For Debug Only!

End Function

' <-------------------------------------------------------------------------------------------------->
'
'     Module Name:  ICL.SetErrorInformation (Private)
'
'     Description:  Formats the error description for line numbering and any addlitional information
'                   that should be added.
'
'          Inputs:  vlngErrNumber, the error number
'                   vstrErrSource, the error source
'                   vstrErrDescr, the original error description.
'                   vlngErrLine, the line number the error occurred at.
'                   rudtErrorInfo, a reference to the error information type to be updated
'
'         Outputs:  rudtErrorInfo, updated with error number, source, formatted error description including last SQL statement
'
'  Change History:
'
'           Initials+#  Reference   Tracking #  Date        Description
'           ----------  ---------   ----------  ----------  ----------------------------------------------------
'           rda001      ICL         n/a         05/15/2003  Created for Project
'
Private Sub SetErrorInformation(ByVal vlngErrNumber As Long, ByVal vstrErrSource As String, ByVal vstrErrDescr As String, _
                                ByVal vlngErrLine As Long, ByRef rudtErrorInfo As mudtErrorData)

    On Error Resume Next

    ' Update the Error type with the error information
    With rudtErrorInfo
        ' If an exit without error was raised
            
        ' Save the error information
        .Number = vlngErrNumber
        .Source = vstrErrSource
        .LineNumber = vlngErrLine
        
        ' Build the extended error description
        
        ' Set the error description to the procedure's name.
        .Desc = .Method
        
        ' Check to see if the line number is blank or is set to "0", if not, the add it to the
        ' description with the proper preface.
        If (vlngErrLine <> 0) Then
            .Desc = .Desc & " at line: " & vlngErrLine
        End If
        
        ' Check to see if the additional information is blank. If not, then surround it with parenthesis
        ' and add it to the error description at the end.
        If (Len(.AddlInfo) <> 0) Then
            .Desc = .Desc & " (" & .AddlInfo & ")"
        End If
        
        ' Add an error seperator character ">" and append the original error description.
        .Desc = .Desc & " > " & vstrErrDescr
    
    End With

End Sub

' <-------------------------------------------------------------------------------------------------->
'
'     Module Name:  ICL.LastSystemError (Private)
'
'     Description:  Returns the message from the system which describes the
'                   last dll error to occur, as held in Err.LastDllError. This function should be
'                   called as soon after the API call which might have errored, as this member can be reset
'                   to zero by subsequent API calls.
'
'     Pseudo Code:  (None)
'
'          Inputs:  (None)
'
'         Outputs:  String, containing data from the O/S pertaining to the last system error.
'
'  Change History:
'
'           Initials+#  Reference   Tracking #  Date        Description
'           ----------  ---------   ----------  ----------  ----------------------------------------------------
'           rda001      ICL         n/a         05/15/2003  Created for Project
'
Private Function LastSystemError() As String

Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

Dim strError    As String * 500 ' Preinitilize a string buffer to put any error message into
Dim lngErrNum   As Long
Dim lngErrMsg   As Long

    lngErrNum = Err.LastDllError
    lngErrMsg = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, lngErrNum, 0, strError, Len(strError), 0)
    LastSystemError = Trim$(strError)

End Function

' <-------------------------------------------------------------------------------------------------->
'
'     Module Name:  ICL.WidthBytes (Private)
'
'     Description:  MACRO: #define WIDTHBYTES(bits)      ((((bits) + 31)>>5)<<2)
'                   Returns the width of a row in a DIB Bitmap given the
'                   number of bits.  DIB Bitmap rows always align on a DWORD boundary.
'
'     Pseudo Code:  (None)
'
'          Inputs:  vlngWidth, ByVal Long
'
'         Outputs:  Long
'
'  Change History:
'
'           Initials+#  Reference   Tracking #  Date        Description
'           ----------  ---------   ----------  ----------  ----------------------------------------------------
'           rda001      ICL         n/a         05/15/2003  Created for Project
'
Private Function WidthBytes(ByVal vlngWidth As Long) As Long
   WidthBytes = ((vlngWidth + 31) \ 32) * 4
End Function

' <-------------------------------------------------------------------------------------------------->
'
'     Module Name:  ICL.ReadLibrary (Public)
'
'     Description:  Opens the ICL library file and imports all the Icons data therein.
'
'     Pseudo Code:  (None)
'
'          Inputs:  (None)
'
'         Outputs:  (None)
'
'  Change History:
'
'           Initials+#  Reference   Tracking #  Date        Description
'           ----------  ---------   ----------  ----------  ----------------------------------------------------
'           rda001      ICL         n/a         05/15/2003  Created for Project
'
Public Sub ReadLibrary()

    On Error GoTo ErrorCode

Dim udtErrInfo          As mudtErrorData
Dim udtDosStub          As IMAGE_DOS_HEADER
Dim udtNEHeader         As IMAGE_WINDOWS_NE_HEADER
Dim udtIconGroupData    As IconGroupData
Dim udtMemIcon          As MEMICONDIRENTRY
Dim udtIconDir          As ICONDIR
Dim udtBitmap           As IconBitmapData

Dim lngFileHandle       As Long
Dim lngStartOffset      As Long
Dim lngBytesRead        As Long
Dim lngAlignShift       As Long
Dim lngOffset           As Long
Dim lngTypeID           As Long
Dim lngResCount         As Long
Dim lngCounter          As Long
Dim lng2ndCounter       As Long
Dim lngTemp             As Long
Dim lngColors           As Long

Dim audtIconInfo()      As TNAMEINFO
Dim audtIconGroupInfo() As TNAMEINFO
Dim audtIconGroupData() As IconGroupData
Dim astrNames()         As String
Dim abytData()          As Byte
Dim audtBitmaps()       As IconBitmapData

Dim lngLength           As Long
Dim lngBitmapLength     As Long

    udtErrInfo.Method = App.Title & mcstrMODULE_BASE_NAME & "ReadLibrary()"
    
    lngFileHandle = OpenFile(mstrFileName, OpenMode.ForReading)
    
    ' Read in the DOS Stub and jump to the start of the Windows (NE) Header
    Call SetFilePointer(lngFileHandle, 0, 0, FILE_BEGIN)
    If Err.LastDllError Then
        Err.Raise menmErrorNumber.ERR_READING_FILE, udtErrInfo.Method, LastSystemError()
    End If
    
    Call ReadFile(lngFileHandle, VarPtr(udtDosStub), Len(udtDosStub), lngBytesRead, ByVal 0&)
    If Err.LastDllError Then
        Err.Raise menmErrorNumber.ERR_READING_FILE, udtErrInfo.Method, LastSystemError()
    End If
    
    lngStartOffset = udtDosStub.e_lfanew
    Call SetFilePointer(lngFileHandle, lngStartOffset, 0, FILE_BEGIN)
       
    ' Read in the Windows (NE) Header
    Call ReadFile(lngFileHandle, VarPtr(udtNEHeader), Len(udtNEHeader), lngBytesRead, ByVal 0&)
    If Err.LastDllError Then
        udtErrInfo.AddlInfo = "Reading Windows Header"
        Err.Raise menmErrorNumber.ERR_READING_FILE, udtErrInfo.Method, LastSystemError()
    End If
    
    ' Move to the Resource table.
    Call SetFilePointer(lngFileHandle, lngStartOffset + udtNEHeader.ResourceTableOffset, 0, FILE_BEGIN)
        
    ' Begin Reading the table
    ' rscAlignShift
    Call ReadFile(lngFileHandle, VarPtr(lngAlignShift), 2, lngBytesRead, ByVal 0&)
    If Err.LastDllError Then
        Err.Raise menmErrorNumber.ERR_READING_FILE, udtErrInfo.Method, LastSystemError()
    End If
    lngAlignShift = 2 ^ lngAlignShift
    
    ' rscTypes
    Do
    
        ' rtTypeID
        Call ReadFile(lngFileHandle, VarPtr(lngTypeID), 2, lngBytesRead, ByVal 0&)
        If Err.LastDllError Then
            Err.Raise menmErrorNumber.ERR_READING_FILE, udtErrInfo.Method, LastSystemError()
        End If
        If (lngTypeID <= 0) Then Exit Do
        If (lngTypeID > RT_HIGHORDER) Then
            lngTypeID = lngTypeID - RT_HIGHORDER
        End If
        
        ' rtResourceCount
        Call ReadFile(lngFileHandle, VarPtr(lngResCount), 2, lngBytesRead, ByVal 0&)
        If Err.LastDllError Then
            Err.Raise menmErrorNumber.ERR_READING_FILE, udtErrInfo.Method, LastSystemError()
        End If
        
        ' rtReserved
        Call ReadFile(lngFileHandle, VarPtr(lngTemp), 4, lngBytesRead, ByVal 0&)
        If Err.LastDllError Then
            Err.Raise menmErrorNumber.ERR_READING_FILE, udtErrInfo.Method, LastSystemError()
        End If
        
        ' rtNameInfo
        Select Case lngTypeID
        
            Case RT_GROUP_ICON
                ReDim audtIconGroupInfo(lngResCount - 1)
                ReDim audtIconGroupData(lngResCount - 1)
                Call ReadFile(lngFileHandle, VarPtr(audtIconGroupInfo(0)), Len(audtIconGroupInfo(0)) * lngResCount, lngBytesRead, ByVal 0&)
                        
            Case RT_ICON
                ReDim audtIconInfo(lngResCount - 1)
                ReDim audtBitmaps(lngResCount - 1)
                Call ReadFile(lngFileHandle, VarPtr(audtIconInfo(0)), Len(audtIconInfo(0)) * lngResCount, lngBytesRead, ByVal 0&)
                
        End Select
                
        If Err.LastDllError Then
            Err.Raise menmErrorNumber.ERR_READING_FILE, udtErrInfo.Method, LastSystemError()
        End If
    
    Loop Until lngTypeID <= 0
    
    lngCounter = -1
    
    ' rscResourceNames
    Do
    
        Call ReadFile(lngFileHandle, VarPtr(lngLength), 1, lngBytesRead, ByVal 0&)
        If Err.LastDllError Then
            Err.Raise menmErrorNumber.ERR_READING_FILE, udtErrInfo.Method, LastSystemError()
        End If
        
        If (lngLength <= 0) Then Exit Do
        
        ReDim abytData(lngLength - 1)
        Call ReadFile(lngFileHandle, VarPtr(abytData(0)), lngLength, lngBytesRead, ByVal 0&)
        
        lngCounter = lngCounter + 1
        ReDim Preserve astrNames(lngCounter)
        
        astrNames(lngCounter) = StrConv(abytData(), vbUnicode, 1033)
        
    Loop Until lngLength <= 0
    Erase abytData()
    
    ' Now we are in the resource data inself. Read in each of the IconGroup headers
    For lngCounter = 0 To UBound(audtIconGroupData())
    
        lngOffset = audtIconGroupInfo(lngCounter).rnOffset
        If (lngOffset < 0) Then
            lngOffset = lngOffset + 65536 ' UWORD
        End If
        lngOffset = lngOffset * lngAlignShift
       
        lngLength = audtIconGroupInfo(lngCounter).rnLength
        If (lngLength < 0) Then
            lngLength = lngLength + 65536 ' UWORD
        End If
        lngLength = lngLength * lngAlignShift
        
        Call SetFilePointer(lngFileHandle, lngOffset, 0, FILE_BEGIN)
        
        Call ReadFile(lngFileHandle, VarPtr(udtIconDir), Len(udtIconDir), lngBytesRead, ByVal 0&)
        audtIconGroupData(lngCounter).DirHeader = udtIconDir
        
        ReDim audtIconGroupData(lngCounter).Entries(udtIconDir.idCount - 1)
        For lng2ndCounter = 0 To udtIconDir.idCount - 1
            Call ReadFile(lngFileHandle, VarPtr(udtMemIcon), Len(udtMemIcon), lngBytesRead, ByVal 0&)
            audtIconGroupData(lngCounter).Entries(lng2ndCounter) = udtMemIcon
        Next
        
    Next
    
    ' Now load the bitmap data of the icons
    For lngCounter = 0 To UBound(audtIconInfo())
    
        lngOffset = audtIconInfo(lngCounter).rnOffset
        
        If (lngOffset < 0) Then
            lngOffset = lngOffset + 65536 ' UWORD
        End If
        lngOffset = lngOffset * lngAlignShift
        
        
        lngBitmapLength = audtIconInfo(lngCounter).rnLength
        If (lngBitmapLength < 0) Then
            lngBitmapLength = lngBitmapLength + 65536 ' UWORD
        End If
        lngBitmapLength = lngBitmapLength * lngAlignShift
        
        Call SetFilePointer(lngFileHandle, lngOffset, 0, FILE_BEGIN)
        
        ' Bitmap Header
        Call ReadFile(lngFileHandle, VarPtr(udtBitmap.Header), Len(udtBitmap.Header), lngBytesRead, ByVal 0&)
        lngBitmapLength = lngBitmapLength - Len(udtBitmap.Header)
        
        ' Palette
        lngColors = GetColorCount(udtBitmap.Header.biBitCount)
        If (lngColors > 0) Then
            ReDim udtBitmap.Palette(lngColors - 1)
            Call ReadFile(lngFileHandle, VarPtr(udtBitmap.Palette(0)), Len(udtBitmap.Palette(0)) * lngColors, lngBytesRead, ByVal 0&)
            lngBitmapLength = lngBitmapLength - Len(udtBitmap.Palette(0)) * lngColors
        Else
            Erase udtBitmap.Palette()
        End If

        ' XOR Map (Color Channel)
        lngLength = udtBitmap.Header.biHeight / 2 * WidthBytes(udtBitmap.Header.biWidth * udtBitmap.Header.biBitCount * udtBitmap.Header.biPlanes)
        ReDim udtBitmap.XORMap(lngLength - 1)
        Call ReadFile(lngFileHandle, VarPtr(udtBitmap.XORMap(0)), lngLength, lngBytesRead, ByVal 0&)
        lngBitmapLength = lngBitmapLength - lngLength
            
        ' AND Map (Alpha Channel)
        ReDim udtBitmap.ANDMap(lngBitmapLength - 1)
        Call ReadFile(lngFileHandle, VarPtr(udtBitmap.ANDMap(0)), lngBitmapLength, lngBytesRead, ByVal 0&)
        
        audtBitmaps(lngCounter) = udtBitmap
        
    Next
    
    Call FlushFileBuffers(lngFileHandle)
    Call CloseHandle(lngFileHandle)
    lngFileHandle = 0
    
    ' Now that we've loaded up all the data, reorganize it properly.
    lngResCount = UBound(audtIconGroupData())
    ReDim maudtIconData(lngResCount)
    For lngCounter = 0 To lngResCount
        With maudtIconData(lngCounter)
            .Name = astrNames(lngCounter + 1)
            .Group = audtIconGroupData(lngCounter)
            ReDim .Bitmaps(.Group.DirHeader.idCount - 1)
            For lng2ndCounter = 0 To .Group.DirHeader.idCount - 1
                .Bitmaps(lng2ndCounter) = audtBitmaps(.Group.Entries(lng2ndCounter).nID - 1)
            Next
        End With
    Next
    
    mlngIconCount = lngResCount + 1
    
' <---------------------------------------------------------->
ExitCode:

    ' Deallocate all object references.
    If (lngFileHandle <> 0) Then
        Call FlushFileBuffers(lngFileHandle)
        Call CloseHandle(lngFileHandle)
    End If
    
    Erase audtIconInfo()
    Erase audtIconGroupInfo()
    Erase audtIconGroupData()
    Erase astrNames()
    Erase abytData()
    Erase audtBitmaps()

    ' If an error occurs, raise the Err object and exit the function.
    If (udtErrInfo.Number <> menmErrorNumber.No_Error) Then
        Err.Raise udtErrInfo.Number, udtErrInfo.Source, udtErrInfo.Desc
    End If

    Exit Sub

' <---------------------------------------------------------->
ErrorCode:

    ' Set a permanent break point for the IDE. This has no effect in compiled code.
    Debug.Assert False

    ' Error occurred store it, stop error catching, and go to exit
    Call SetErrorInformation(Err.Number, Err.Source, Err.Description, Erl(), udtErrInfo)
    On Error GoTo 0
    Resume ExitCode
    Resume ' For Debug Only!

End Sub

' <---------------------------------------------------------------------------------------->
'
'     Module Name:  ICL.DeleteIcon (Private)
'
'     Description:  Deletes an icon from the array of Icon data based on the Index given.
'
'      Dependency:  (None)
'
'          Inputs:  vlngIndex, ByVal Long
'
'         Outputs:  Long
'
'  Change History:
'
'           Initials+#  Reference   Tracking #  Date        Description
'           ----------  ----------- ----------  ----------  ----------------------
'           rda001      ICL         n/a         05/23/2003  Created for Project
'
Public Function DeleteIcon(ByVal vlngIndex As Long) As Long

    On Error GoTo ErrorCode

Dim udtErrInfo  As mudtErrorData
Dim lngCounter  As Long
Dim lngTotal    As Long

    udtErrInfo.Method = App.Title & mcstrMODULE_BASE_NAME & "DeleteIcon()"

    ' Make sure the Index passed in is valid.
    If ((vlngIndex < 1) Or (vlngIndex > mlngIconCount)) Then
        udtErrInfo.AddlInfo = CStr(vlngIndex)
        Err.Raise menmErrorNumber.ERR_DATA_OUT_OF_RANGE, udtErrInfo.Method, "Invalid Index."
    End If

    ' Move all the remaining Icon data up one notch in the array.
    lngTotal = (mlngIconCount - 1) - 1
    For lngCounter = vlngIndex To lngTotal
        maudtIconData(lngCounter) = maudtIconData(lngCounter + 1)
    Next

    mlngIconCount = (mlngIconCount - 1) - 1
    ReDim Preserve maudtIconData(mlngIconCount)

    Call ReorderIconIDs

' <---------------------------------------------------------->
ExitCode:

    ' Deallocate all object references.
    ' If an error occurs, raise the Err object and exit the function.
    If (udtErrInfo.Number <> menmErrorNumber.No_Error) Then
        Err.Raise udtErrInfo.Number, udtErrInfo.Source, udtErrInfo.Desc, udtErrInfo.HelpFile, udtErrInfo.HelpContext
    End If

    Exit Function

' <---------------------------------------------------------->
ErrorCode:

    ' Set a permanent break point for the IDE. This has no effect in compiled code.
    Debug.Assert False

    ' Error occurred store it, stop error catching, and go to exit
    Call SetErrorInformation(Err.Number, Err.Source, Err.Description, Erl(), udtErrInfo)
    On Error GoTo 0
    Resume ExitCode
    Resume ' For Debug Only!

End Function

' <---------------------------------------------------------------------------------------->
'
'     Module Name:  ICL.AddIcon (Private)
'
'     Description:  Adds an Icon from file to the Icon Library. Returns the Index of the new Icon (Group).
'
'      Dependency:  (None)
'
'          Inputs:  vstrFilename, ByVal String
'
'         Outputs:  Long
'
'  Change History:
'
'           Initials+#  Reference   Tracking #  Date        Description
'           ----------  ----------- ----------  ----------  ----------------------
'           rda001      ICL         n/a         05/22/2003  Created for Project
'
Public Function AddIcon(ByVal vstrFilename As String) As Long

    On Error GoTo ErrorCode

Dim udtErrInfo      As mudtErrorData
Dim udtIconData     As IconData
Dim audtEntries()   As ICONDIRENTRY
Dim udtBitmap       As IconBitmapData
Dim hFile           As Long
Dim lngTemp         As Long
Dim lngTotal        As Long
Dim lngCounter      As Long
Dim lngBytesRead    As Long
Dim lngBitmapLength As Long
Dim lngColors       As Long
Dim lngLength       As Long
Dim strName         As String

    udtErrInfo.Method = App.Title & mcstrMODULE_BASE_NAME & "AddIcon()"
    
    hFile = OpenFile(vstrFilename, OpenMode.ForReading)
    Call SetFilePointer(hFile, 0, 0, FILE_BEGIN)
    If Err.LastDllError Then
        udtErrInfo.AddlInfo = "Unable to open Icon file: """ & vstrFilename & """"
        Err.Raise menmErrorNumber.ERR_READING_FILE, udtErrInfo.Method, LastSystemError()
    End If
    
    ' Begin Reading the File
    ' Reserved (WORD)
    Call ReadFile(hFile, VarPtr(lngTemp), 2, lngBytesRead, ByVal 0&)
        
    ' ResourceType (WORD) Always = 1
    Call ReadFile(hFile, VarPtr(lngTemp), 2, lngBytesRead, ByVal 0&)
    If (lngTemp <> 1) Then
        udtErrInfo.AddlInfo = CStr(lngTemp)
        Err.Raise menmErrorNumber.ERR_INVALID_DATA_FORMAT, udtErrInfo.Method, "Invalid Resource Type encountered."
    End If
    
    ' IconCount (WORD)
    Call ReadFile(hFile, VarPtr(lngTotal), 2, lngBytesRead, ByVal 0&)
    If (lngTotal <= 0) Then
        Err.Raise menmErrorNumber.ERR_DATA_OUT_OF_RANGE, udtErrInfo.Method, "No Icons in file."
    End If
    
    ' IconDir[]
    ReDim audtEntries(lngTotal - 1)
    ReDim udtIconData.Bitmaps(lngTotal - 1)
    
    Call ReadFile(hFile, VarPtr(audtEntries(0)), Len(audtEntries(0)) * lngTotal, lngBytesRead, ByVal 0&)
    
    udtIconData.Group.DirHeader.idCount = lngTotal
    
    ReDim udtIconData.Group.Entries(lngTotal - 1)
    
    ' IconData[] (Bitmap data)
    For lngCounter = 0 To lngTotal - 1
    
        Call SetFilePointer(hFile, audtEntries(lngCounter).DataOffset, 0, FILE_BEGIN)
        lngBitmapLength = audtEntries(lngCounter).DataSize
    
        With udtIconData.Group.Entries(lngCounter)
            .bColorCount = audtEntries(lngCounter).ColorCount
            .bHeight = audtEntries(lngCounter).Height
            .bWidth = audtEntries(lngCounter).Width
            .wBitCount = audtEntries(lngCounter).BitsPerPixel
            .dwBytesInRes = lngBitmapLength
        End With
    
        Call ReadFile(hFile, VarPtr(udtBitmap.Header), Len(udtBitmap.Header), lngBytesRead, ByVal 0&)
        lngBitmapLength = lngBitmapLength - Len(udtBitmap.Header)
        
        ' Palette
        lngColors = GetColorCount(udtBitmap.Header.biBitCount)
        If (lngColors > 0) Then
            ReDim udtBitmap.Palette(lngColors - 1)
            Call ReadFile(hFile, VarPtr(udtBitmap.Palette(0)), Len(udtBitmap.Palette(0)) * lngColors, lngBytesRead, ByVal 0&)
            lngBitmapLength = lngBitmapLength - Len(udtBitmap.Palette(0)) * lngColors
        Else
            Erase udtBitmap.Palette()
        End If

        ' XOR Map (Color Channel)
        lngLength = udtBitmap.Header.biHeight / 2 * WidthBytes(udtBitmap.Header.biWidth * udtBitmap.Header.biBitCount * udtBitmap.Header.biPlanes)
        ReDim udtBitmap.XORMap(lngLength - 1)
        Call ReadFile(hFile, VarPtr(udtBitmap.XORMap(0)), lngLength, lngBytesRead, ByVal 0&)
        lngBitmapLength = lngBitmapLength - lngLength
            
        ' AND Map (Alpha Channel)
        ReDim udtBitmap.ANDMap(lngBitmapLength - 1)
        Call ReadFile(hFile, VarPtr(udtBitmap.ANDMap(0)), lngBitmapLength, lngBytesRead, ByVal 0&)
        
        udtIconData.Bitmaps(lngCounter) = udtBitmap
        
    Next
    
    ' strip off the path and the extension of the filename
    strName = Mid$(Left$(vstrFilename, InStr(vstrFilename, ".") - 1), InStrRev(vstrFilename, "\") + 1)
    If (Len(strName) > 0) Then
        Mid$(strName, 1, 1) = UCase$(Mid$(strName, 1, 1))
    Else
        strName = "Icon"
    End If
    udtIconData.Name = strName
        
    mlngIconCount = mlngIconCount + 1
    ReDim Preserve maudtIconData(mlngIconCount - 1)
    maudtIconData(mlngIconCount - 1) = udtIconData
    
    Call ReorderIconIDs
    
    AddIcon = mlngIconCount
    
' <---------------------------------------------------------->
ExitCode:

    ' Flush the buffers and close the file
    If (hFile <> 0) Then
        Call FlushFileBuffers(hFile)
        Call CloseHandle(hFile)
    End If

    ' Deallocate all object references.
    Erase audtEntries()
    
    ' If an error occurs, raise the Err object and exit the function.
    If (udtErrInfo.Number <> menmErrorNumber.No_Error) Then
        Err.Raise udtErrInfo.Number, udtErrInfo.Source, udtErrInfo.Desc, udtErrInfo.HelpFile, udtErrInfo.HelpContext
    End If

    Exit Function

' <---------------------------------------------------------->
ErrorCode:

    ' Set a permanent break point for the IDE. This has no effect in compiled code.
    Debug.Assert False
    
    ' Error occurred store it, stop error catching, and go to exit
    Call SetErrorInformation(Err.Number, Err.Source, Err.Description, Erl(), udtErrInfo)
    On Error GoTo 0
    Resume ExitCode
    Resume ' For Debug Only!

End Function

' <---------------------------------------------------------------------------------------->
'
'     Module Name:  ICL.ExportIcon (Public)
'
'     Description:  Constructs an Icon file and exports it to disk.
'
'      Dependency:  (None)
'
'          Inputs:  vlngIndex, ByVal Long, Index of the Icon to export
'                   vstrFilename, ByVal String, name of the file to write.
'
'         Outputs:  (None)
'
'  Change History:
'
'           Initials+#  Reference   Tracking #  Date        Description
'           ----------  ----------- ----------  ----------  ----------------------
'           rda001      ICL         n/a         05/23/2003  Created for Project
'
Public Sub ExportIcon(ByVal vlngIndex As Long, ByVal vstrFilename As String)

    On Error GoTo ErrorCode

Dim udtErrInfo          As mudtErrorData
Dim udtIconData         As IconData
Dim udtIconEntry        As IconEntry
Dim udtBitmap           As IconBitmapData
Dim alngOffsets()       As Long
Dim lngCounter          As Long
Dim lngTotal            As Long
Dim hFile               As Long ' File Handle
Dim lngBytesWritten     As Long
Dim lngLength           As Long
Dim abytData()          As Byte
Dim lngColors           As Long

    udtErrInfo.Method = App.Title & mcstrMODULE_BASE_NAME & "ExportIcon()"

    ' Make sure the Index passed in is valid.
    If ((vlngIndex < 1) Or (vlngIndex > mlngIconCount)) Then
        udtErrInfo.AddlInfo = CStr(vlngIndex)
        Err.Raise menmErrorNumber.ERR_DATA_OUT_OF_RANGE, udtErrInfo.Method, "Invalid Index."
    End If

    udtIconData = maudtIconData(vlngIndex - 1)
    hFile = OpenFile(vstrFilename, OpenMode.ForWriting)

    ' ------------------------------------------------------------
    ' Go to start of file...
    Call SetFilePointer(hFile, 0, 0, FILE_BEGIN)
    If Err.LastDllError Then
        udtErrInfo.AddlInfo = "Resetting file pointer to start of file"
        Err.Raise menmErrorNumber.ERR_WRITING_FILE, udtErrInfo.Method, LastSystemError()
    End If

    ' Reserved (WORD)
    Call WriteFile(hFile, 0&, 2, lngBytesWritten, &H0)
    
    ' Resource Type (WORD) = 1
    Call WriteFile(hFile, 1&, 2, lngBytesWritten, &H0)

    ' Icon Count (WORD)
    lngTotal = udtIconData.Group.DirHeader.idCount
    Call WriteFile(hFile, lngTotal, 2, lngBytesWritten, &H0)
    
    udtErrInfo.AddlInfo = "Writing Icon Directory Info"
    
    ' Icon Directory Array
    lngLength = 16 ' Len(udtIconEntry)
    ReDim abytData(lngLength - 1)

    For lngCounter = 0 To lngTotal - 1

        With udtIconData.Group.Entries(lngCounter)
            udtIconEntry.Width = .bWidth
            udtIconEntry.Height = .bHeight
            udtIconEntry.NumColors = .bColorCount
            udtIconEntry.Reserved = 0
            udtIconEntry.NumPlanes = .wPlanes
            udtIconEntry.BitsPerPixel = .wBitCount
            udtIconEntry.DataSize = .dwBytesInRes
            udtIconEntry.DataOffset = 0&    ' Use a zero placeholder for now
        End With

        Call CopyMemory(abytData(0), udtIconEntry, lngLength)
        Call WriteFile(hFile, abytData(0), lngLength, lngBytesWritten, &H0)
        If Err.LastDllError Then
            Err.Raise menmErrorNumber.ERR_WRITING_FILE, udtErrInfo.Method, LastSystemError()
        End If

    Next

    udtErrInfo.AddlInfo = "Writing Icon Bitmap Info"
    
    ' Icon Data Array
    ReDim alngOffsets(lngTotal - 1)
    For lngCounter = 0 To lngTotal - 1

        ' Get the current offset value.
        alngOffsets(lngCounter) = SetFilePointer(hFile, 0, 0, FILE_CURRENT)

        udtBitmap = udtIconData.Bitmaps(lngCounter)

        ' Header
        lngLength = Len(udtBitmap.Header)
        ReDim abytData(lngLength - 1)
        Call CopyMemory(abytData(0), udtBitmap.Header, lngLength)
        Call WriteFile(hFile, abytData(0), lngLength, lngBytesWritten, &H0)
        If Err.LastDllError Then
            Err.Raise menmErrorNumber.ERR_WRITING_FILE, udtErrInfo.Method, LastSystemError()
        End If

        ' Palette
        lngColors = GetColorCount(udtBitmap.Header.biBitCount)
        If (lngColors > 0) Then
            Call WriteFile(hFile, udtBitmap.Palette(0), Len(udtBitmap.Palette(0)) * lngColors, lngBytesWritten, &H0)
            If Err.LastDllError Then
                Err.Raise menmErrorNumber.ERR_WRITING_FILE, udtErrInfo.Method, LastSystemError()
            End If
        End If

        ' XOR Map (Color Channel)
        lngLength = UpperBounds(udtBitmap.XORMap()) + 1
        Call WriteFile(hFile, udtBitmap.XORMap(0), lngLength, lngBytesWritten, &H0)
        If Err.LastDllError Then
            Err.Raise menmErrorNumber.ERR_WRITING_FILE, udtErrInfo.Method, LastSystemError()
        End If

        ' AND Map (Alpha Channel)
        lngLength = UpperBounds(udtBitmap.ANDMap()) + 1
        Call WriteFile(hFile, udtBitmap.ANDMap(0), lngLength, lngBytesWritten, &H0)
        If Err.LastDllError Then
            Err.Raise menmErrorNumber.ERR_WRITING_FILE, udtErrInfo.Method, LastSystemError()
        End If

    Next

    ' Now go back and update the offsets.
    udtErrInfo.AddlInfo = "Writing Offset Values"
    Call SetFilePointer(hFile, 18, 0, FILE_BEGIN)
    
    For lngCounter = 0 To lngTotal - 1
        Call WriteFile(hFile, alngOffsets(lngCounter), 2, lngBytesWritten, &H0)
        If Err.LastDllError Then
            Err.Raise menmErrorNumber.ERR_WRITING_FILE, udtErrInfo.Method, LastSystemError()
        End If
        Call SetFilePointer(hFile, 14, 0, FILE_CURRENT)

    Next

' <---------------------------------------------------------->
ExitCode:

    ' Flush the buffers and close the file
    If (hFile <> 0) Then
        Call FlushFileBuffers(hFile)
        Call CloseHandle(hFile)
    End If

    ' Deallocate all object references.
    Erase abytData()
    Erase alngOffsets()
    
    ' If an error occurs, raise the Err object and exit the function.
    If (udtErrInfo.Number <> menmErrorNumber.No_Error) Then
        Err.Raise udtErrInfo.Number, udtErrInfo.Source, udtErrInfo.Desc, udtErrInfo.HelpFile, udtErrInfo.HelpContext
    End If

    Exit Sub

' <---------------------------------------------------------->
ErrorCode:

    ' Set a permanent break point for the IDE. This has no effect in compiled code.
    Debug.Assert False

    ' Error occurred store it, stop error catching, and go to exit
    Call SetErrorInformation(Err.Number, Err.Source, Err.Description, Erl(), udtErrInfo)
    On Error GoTo 0
    Resume ExitCode
    Resume ' For Debug Only!

End Sub

' <---------------------------------------------------------------------------------------->
'
'     Module Name:  ICL.Init (Public)
'
'     Description:  Initialize then property variables.
'
'      Dependancy:  (None)
'
'          Inputs:  (None)
'
'         Outputs:  (None)
'
'  Change History:
'
'           Initials+#  Reference   Tracking #  Date        Description
'           ----------  ---------   ----------  ----------  ----------------------------------------------------
'           rda001      ICL         n/a         05/15/2003  Created for Project
'
Public Sub Init()

    On Error GoTo ErrorCode

Dim udtErrInfo  As mudtErrorData

    udtErrInfo.Method = App.Title & mcstrMODULE_BASE_NAME & "Init()"

    mstrFileName = vbNullString
    Erase maudtIconData()
    
' <---------------------------------------------------------->
ExitCode:

    ' Deallocate all object references.

    ' If an error occurs, raise the Err object and exit the function.
    If (udtErrInfo.Number <> menmErrorNumber.No_Error) Then
        Err.Raise udtErrInfo.Number, udtErrInfo.Source, udtErrInfo.Desc
    End If

    Exit Sub

' <---------------------------------------------------------->
ErrorCode:

            ' Set a permanent break point for the IDE. This has no effect in compiled code.
    Debug.Assert False

    ' Error occurred store it, stop error catching, and go to exit
    Call SetErrorInformation(Err.Number, Err.Source, Err.Description, Erl(), udtErrInfo)
    On Error GoTo 0
    Resume ExitCode
    Resume ' For Debug Only!

End Sub

' <---------------------------------------------------------------------------------------->
'
'     Module Name:  ICL.Term (Public)
'
'     Description:  Perform general cleanup within the class before it is destroyed:
'                   deallocate the object(s) used in this class as properties and erase
'                   any array(s).
'
'      Dependancy:  (None)
'
'          Inputs:  (None)
'
'         Outputs:  (None)
'
'  Change History:
'
'           Initials+#  Reference   Tracking #  Date        Description
'           ----------  ---------   ----------  ----------  ----------------------------------------------------
'           rda001      ICL         n/a         05/15/2003  Created for Project
'
Public Sub Term()

    On Error GoTo ErrorCode

Dim udtErrInfo As mudtErrorData

    udtErrInfo.Method = App.Title & mcstrMODULE_BASE_NAME & "Term()"

    ' Clear all module level variables here!
    mstrFileName = vbNullString
    
    ' Destroy all module level objects here!
    
    ' Empty all module level arrays here!
    Erase maudtIconData()
    
' <---------------------------------------------------------->
ExitCode:

    ' Deallocate all object references.

    ' If an error occurs, raise the Err object and exit the function.
    If (udtErrInfo.Number <> menmErrorNumber.No_Error) Then
        Err.Raise udtErrInfo.Number, udtErrInfo.Source, udtErrInfo.Desc
    End If

    Exit Sub

' <---------------------------------------------------------->
ErrorCode:

    ' Error occurred store it, stop error catching, and go to exit
    Call SetErrorInformation(Err.Number, Err.Source, Err.Description, Erl(), udtErrInfo)
    On Error GoTo 0
    Resume ExitCode
    Resume ' For Debug Only!

End Sub


' <-------------------------------------------------------------------------------------------------->
'
'     Module Name:  ICL.WriteLibrary (Public)
'
'     Description:  Opens a new or existing file and writes out the Icon Library Data.
'
'     Pseudo Code:  (None)
'
'          Inputs:  vstrFileName, ByVal String
'
'         Outputs:  (None)
'
'  Change History:
'
'           Initials+#  Reference   Tracking #  Date        Description
'           ----------  ---------   ----------  ----------  ----------------------------------------------------
'           rda001      ICL         n/a         05/15/2003  Created for Project
'
Public Sub WriteLibrary(ByVal vstrFilename As String)

    On Error GoTo ErrorCode

Dim udtErrInfo          As mudtErrorData

Dim udtDosStub          As IMAGE_DOS_HEADER
Dim udtNEHeader         As IMAGE_WINDOWS_NE_HEADER
Dim udtTNameInfo        As TNAMEINFO
Dim udtIconDir          As ICONDIR
Dim udtBitmap           As IconBitmapData
Dim udtMemIcon          As MEMICONDIRENTRY
Dim audtIconGroupDOL()  As DataOffsetLength
Dim audtIconDOL()       As DataOffsetLength

Dim hFile               As Long ' File Handle
Dim lngLength           As Long
Dim lngChunkLength      As Long
Dim lngCounter          As Long
Dim lng2ndCounter       As Long
Dim lngIconGroupCount   As Long ' Number of Icon Groups
Dim lngIconCount        As Long ' Number of Icons
Dim lngBytesWritten     As Long ' Used in Write File
Dim lngIconGroupOffset  As Long
Dim lngIconOffset       As Long
Dim lngOffset           As Long
Dim lngResCount         As Long
Dim lngColors           As Long
Dim lngAlignShift       As Long
Dim strName             As String
Dim dblAlignFactor      As Double

Dim abytData()          As Byte
Dim abytTemp()          As Byte

    udtErrInfo.Method = App.Title & mcstrMODULE_BASE_NAME & "WriteLibrary()"
    
    hFile = OpenFile(vstrFilename, OpenMode.ForWriting)
    
    ' ------------------------------------------------------------
    ' Go to start of file...
    Call SetFilePointer(hFile, 0, 0, FILE_BEGIN)
    If Err.LastDllError Then
        udtErrInfo.AddlInfo = "Resetting file pointer to start of file"
        Err.Raise menmErrorNumber.ERR_WRITING_FILE, udtErrInfo.Method, LastSystemError()
    End If
    
    ' ------------------------------------------------------------
    ' Generate the DOS Stub part of the file.
    udtErrInfo.AddlInfo = "Writing DOS Stub"
    udtDosStub.e_lfanew = 128   ' Location of the start the Windows (NE) Header
    udtDosStub.e_magic = 23117  ' "MZ"
    lngLength = Len(udtDosStub)
    ReDim abytData(lngLength - 1)
    Call CopyMemory(abytData(0), udtDosStub, lngLength)
    Call WriteFile(hFile, abytData(0), lngLength, lngBytesWritten, &H0)
    If Err.LastDllError Then
        Err.Raise menmErrorNumber.ERR_WRITING_FILE, udtErrInfo.Method, LastSystemError()
    End If
    
    ' ------------------------------------------------------------
    ' Write the padding to the Windows (NE) Header
    udtErrInfo.AddlInfo = "Writing DOS Stub padding"
    lngLength = 128 - lngLength
    ReDim abytData(lngLength - 1)
    abytData(lngLength - 2) = &H49 ' "I"
    abytData(lngLength - 1) = &H4D ' "M"
    Call WriteFile(hFile, abytData(0), lngLength, lngBytesWritten, &H0)
    If Err.LastDllError Then
        Err.Raise menmErrorNumber.ERR_WRITING_FILE, udtErrInfo.Method, LastSystemError()
    End If
    
    ' ------------------------------------------------------------
    ' Construct the NE Image area
    udtErrInfo.AddlInfo = "Writing Windows header"
    
    udtNEHeader.Machine = &H454E    ' "NE"
    udtNEHeader.TargetOS = 2
    udtNEHeader.WindowsVersion = &H300
    udtNEHeader.ResourceTableOffset = &H40
    udtNEHeader.SegmentTableOffset = &H40
    
    ' Compute the size of the ResourseTable area to determine the offset to the
    ' Resident Name Table.
    
    ' AlignShift +(TTypeInfo area without the TNameInfo size)
    lngLength = 18 ' 2 + 8 * 2
    
    ' TNameInfo arrays are 12 bytes * Each Resource.
    ' IconGroups are each stored in elements in the array maudtIconData()
    lngIconGroupCount = UBound(maudtIconData()) + 1
    lngLength = lngLength + 12 * lngIconGroupCount
    
    ' Icons are stored within each element in the array maudtIconData(). There can be more
    ' than one icon in each element, so we need to tranverse the array.
    lngIconCount = 0
    For lngCounter = 0 To UBound(maudtIconData())
        lngIconCount = lngIconCount + maudtIconData(lngCounter).Group.DirHeader.idCount
    Next
    lngLength = lngLength + 12 * lngIconCount
    
    ' Add 2 more bytes for the 0 to indicate we stop reading
    lngLength = lngLength + 2
    
    udtNEHeader.ResidentNameTableOffset = &H40 + lngLength
    
    ' Write out the Windows (NE) Header
    lngLength = Len(udtNEHeader)
    ReDim abytData(lngLength - 1)
    Call CopyMemory(abytData(0), udtNEHeader, lngLength)
    Call WriteFile(hFile, abytData(0), lngLength, lngBytesWritten, &H0)
    If Err.LastDllError Then
        Err.Raise menmErrorNumber.ERR_WRITING_FILE, udtErrInfo.Method, LastSystemError()
    End If
        
    ' ------------------------------------------------------------
    ' Now we construct the Resource table.
    
    ' ------------------------------------------------------------
    ' Write out 0 for the AlignShift for now. We will determine how much it needs to be based
    ' the size of the Offsets further on down the method.
    Call WriteFile(hFile, 0&, 2, lngBytesWritten, &H0)
    
    ' Load the data of the Icon Groups first
    ' rtTypeID = &H800E
    Call WriteFile(hFile, RT_HIGHORDER Or RT_GROUP_ICON, 2, lngBytesWritten, &H0)
    
    ' rtResourceCount
    Call WriteFile(hFile, lngIconGroupCount, 2, lngBytesWritten, &H0)
    
    ' Reserved (4 bytes)
    Call WriteFile(hFile, 0&, 4, lngBytesWritten, &H0)
    
    ' ------------------------------------------------------------
    ' Icon Group TTypeInfo (array)
    udtErrInfo.AddlInfo = "Writing TTypeInfo"
    
    ' Get the current file position.
    lngIconGroupOffset = SetFilePointer(hFile, 0, 0, FILE_CURRENT)
    
    udtTNameInfo.rnOffset = 0
    udtTNameInfo.rnLength = 0
    udtTNameInfo.rnFlags = &H1C30
    udtTNameInfo.rnHandle = 0
    udtTNameInfo.rnUsage = 0
    lngLength = Len(udtTNameInfo)
    ReDim abytData(lngLength - 1)
    
    For lngCounter = 0 To lngIconGroupCount - 1
        ' rtNameInfo (array) - All Offsets and lengths are 0'd out for now. They will be populated
        ' after their data is written in and their values are known.
        udtTNameInfo.rnID = CInt("&H" & Hex$(RT_HIGHORDER + (lngCounter + 1)))
        
        ' Move the UDT to the byte array.
        Call CopyMemory(abytData(0), udtTNameInfo, lngLength)
        
        Call WriteFile(hFile, abytData(0), lngLength, lngBytesWritten, &H0)
        If Err.LastDllError Then
            Err.Raise menmErrorNumber.ERR_WRITING_FILE, udtErrInfo.Method, LastSystemError()
        End If
        
    Next

    ' Load the data of the Icons next
    ' rtTypeID = &H8003
    Call WriteFile(hFile, RT_HIGHORDER Or RT_ICON, 2, lngBytesWritten, &H0)
    
    ' rtResourceCount
    Call WriteFile(hFile, lngIconCount, 2, lngBytesWritten, &H0)
    
    ' Reserved (4 bytes)
    Call WriteFile(hFile, 0&, 4, lngBytesWritten, &H0)
    
    ' ------------------------------------------------------------
    ' Icon TTypeInfo (array)
    
    ' Get the current file position.
    lngIconOffset = SetFilePointer(hFile, 0, 0, FILE_CURRENT)
    
    udtTNameInfo.rnFlags = &H1C30
    
    For lngCounter = 0 To lngIconCount - 1
        ' rtNameInfo (array) - All Offsets and lengths are 0'd out for now. They will be populated
        ' after their data is written in and their values are known.
        udtTNameInfo.rnID = CInt("&H" & Hex$(RT_HIGHORDER + (lngCounter + 1)))
        
        ' Move the UDT to the byte array.
        Call CopyMemory(abytData(0), udtTNameInfo, lngLength)
        
        Call WriteFile(hFile, abytData(0), lngLength, lngBytesWritten, &H0)
        If Err.LastDllError Then
            Err.Raise menmErrorNumber.ERR_WRITING_FILE, udtErrInfo.Method, LastSystemError()
        End If
        
    Next
    
    ' Make sure the data is padded to a double word boundary.
    lngOffset = SetFilePointer(hFile, 0, 0, FILE_CURRENT)
    Call PadAmount(hFile, 4, lngOffset)
    
    ' ------------------------------------------------------------
    ' Now write out the Resident Name Table
    ' Each name of the Icon Groups are to be stored as a Pascal string with the length
    ' written out before the string. Move each name and its length to the byte array, then
    ' write the byte array out to the file when done.
    
    udtErrInfo.AddlInfo = "Writing Resident Name Table"
    
    Erase abytData()
    lngOffset = 0
    
    ' Add the string "ICL"
    strName = "ICL"
    Call WriteByteToArray(Len(strName), abytData(), lngOffset)
    Call CopyByteArray(abytData(), StrConv(strName, vbFromUnicode, 1033), lngOffset)

    For lngCounter = 0 To lngIconGroupCount - 1
        strName = maudtIconData(lngCounter).Name
        Call WriteByteToArray(Len(strName), abytData(), lngOffset)
        Call CopyByteArray(abytData(), StrConv(strName, vbFromUnicode, 1033), lngOffset)
    Next

    ' Add a zero to the end to indicate the end of the data
    Call WriteByteToArray(0, abytData(), lngOffset)
    lngLength = UpperBounds(abytData()) + 1
    Call WriteFile(hFile, abytData(0), lngLength, lngBytesWritten, &H0)

    
    ' <-------------------------------------------------------------------->
    ' Estimate the size of the data ahead so we can determine the Size of the AlignShift.
    
    lngChunkLength = 0
    For lngCounter = 0 To lngIconGroupCount - 1
        udtIconDir = maudtIconData(lngCounter).Group.DirHeader
        lngChunkLength = lngChunkLength + 6 ' Group Header Size
        For lng2ndCounter = 0 To udtIconDir.idCount - 1
            lngChunkLength = lngChunkLength + MEMICONDIRENTRY_SIZE  ' Mem Icon Size
            lngChunkLength = lngChunkLength + BITMAP_HEADER_SIZE    ' Bitmap Header Size
            udtBitmap = maudtIconData(lngCounter).Bitmaps(lng2ndCounter)
            lngColors = GetColorCount(udtBitmap.Header.biBitCount)
            If (lngColors > 0) Then
                lngChunkLength = lngChunkLength + Len(udtBitmap.Palette(0)) * lngColors
            End If
            lngChunkLength = lngChunkLength + UpperBounds(udtBitmap.XORMap()) + 1
            lngChunkLength = lngChunkLength + UpperBounds(udtBitmap.ANDMap()) + 1
        Next
    Next
    
    lngChunkLength = lngChunkLength + SetFilePointer(hFile, 0, 0, FILE_CURRENT)
    
    ' Now we can calculate the Align shift and modify all the Offsets and Lengths by that amount.
    dblAlignFactor = lngChunkLength / &HFFFF&
    Select Case dblAlignFactor
        'Case Is < 1#:   lngAlignShift = &H1     ' 2^0
        Case Is < 2#:   lngAlignShift = &H2     ' 2^1
        Case Is < 4#:   lngAlignShift = &H4     ' 2^2
        Case Is < 8#:   lngAlignShift = &H8     ' 2^3
        Case Is < 16#:  lngAlignShift = &H10    ' 2^4
        Case Is < 32#:  lngAlignShift = &H20    ' 2^5
        Case Is < 64#:  lngAlignShift = &H40    ' 2^6
        Case Is < 128#: lngAlignShift = &H80    ' 2^7
    End Select
    
    
    ' Pad to alignment boundary.
    lngOffset = SetFilePointer(hFile, 0, 0, FILE_CURRENT)
    'Call PadAmount(hFile, 4, lngOffset)
    Call PadAmount(hFile, lngAlignShift, lngOffset)
    
    ' <-------------------------------------------------------------------->
    
    udtErrInfo.AddlInfo = "Writing Icon Dir"
        
    ' Record the Current offset position. This is the start of the Icon Group data, and the offset
    ' and length will have to be written back to the file later.
    ReDim audtIconGroupDOL(lngIconGroupCount - 1)
    For lngCounter = 0 To lngIconGroupCount - 1
        
        lngOffset = SetFilePointer(hFile, 0, 0, FILE_CURRENT)
        
        ' Pad to the amount of bytes we are forcing offsets to conform (2, 4, 8...)
        Call PadAmount(hFile, lngAlignShift, lngOffset)
            
        audtIconGroupDOL(lngCounter).offset = lngOffset
    
        udtIconDir = maudtIconData(lngCounter).Group.DirHeader
        lngLength = Len(udtIconDir) ' 22
        lngChunkLength = lngLength
        ReDim abytData(lngLength - 1)
        Call CopyMemory(abytData(0), udtIconDir, lngLength)
        Call WriteFile(hFile, abytData(0), lngLength, lngBytesWritten, &H0)
        If Err.LastDllError Then
            Err.Raise menmErrorNumber.ERR_WRITING_FILE, udtErrInfo.Method, LastSystemError()
        End If
        
        For lng2ndCounter = 0 To udtIconDir.idCount - 1
            udtMemIcon = maudtIconData(lngCounter).Group.Entries(lng2ndCounter)
            lngLength = MEMICONDIRENTRY_SIZE
            lngChunkLength = lngChunkLength + lngLength
            ReDim abytData(lngLength - 1)
            Call CopyMemory(abytData(0), udtMemIcon, lngLength)
            Call WriteFile(hFile, abytData(0), lngLength, lngBytesWritten, &H0)
            If Err.LastDllError Then
                Err.Raise menmErrorNumber.ERR_WRITING_FILE, udtErrInfo.Method, LastSystemError()
            End If
        Next
        
        audtIconGroupDOL(lngCounter).Length = lngChunkLength
        
    Next
    
    ' Now write out the bitmap data to the file.
    udtErrInfo.AddlInfo = "Writing Icon bitmap data"
    
    ' Record the Current offset position. This is the start of the Icon data, and the offset
    ' and length will have to be written back to the file later.
    ReDim audtIconDOL(lngIconCount - 1)
    
    lngResCount = 0
    For lngCounter = 0 To lngIconGroupCount - 1
    
        For lng2ndCounter = 0 To maudtIconData(lngCounter).Group.DirHeader.idCount - 1
        
            lngOffset = SetFilePointer(hFile, 0, 0, FILE_CURRENT)
            Call PadAmount(hFile, lngAlignShift, lngOffset)
            
            audtIconDOL(lngResCount).offset = lngOffset
            udtBitmap = maudtIconData(lngCounter).Bitmaps(lng2ndCounter)
            
            ' Bitmap Header
            lngLength = Len(udtBitmap.Header)
            lngChunkLength = lngLength
            ReDim abytData(lngLength - 1)
            Call CopyMemory(abytData(0), udtBitmap.Header, lngLength)
            Call WriteFile(hFile, abytData(0), lngLength, lngBytesWritten, &H0)
            If Err.LastDllError Then
                Err.Raise menmErrorNumber.ERR_WRITING_FILE, udtErrInfo.Method, LastSystemError()
            End If
    
            ' Palette
            lngColors = GetColorCount(udtBitmap.Header.biBitCount)
            If (lngColors > 0) Then
                lngChunkLength = lngChunkLength + Len(udtBitmap.Palette(0)) * lngColors
                Call WriteFile(hFile, udtBitmap.Palette(0), Len(udtBitmap.Palette(0)) * lngColors, lngBytesWritten, &H0)
                If Err.LastDllError Then
                    Err.Raise menmErrorNumber.ERR_WRITING_FILE, udtErrInfo.Method, LastSystemError()
                End If
            End If
            
            ' XOR Map (Color Channel)
            lngLength = UpperBounds(udtBitmap.XORMap()) + 1
            lngChunkLength = lngChunkLength + lngLength
            Call WriteFile(hFile, udtBitmap.XORMap(0), lngLength, lngBytesWritten, &H0)
            If Err.LastDllError Then
                Err.Raise menmErrorNumber.ERR_WRITING_FILE, udtErrInfo.Method, LastSystemError()
            End If
                
            ' AND Map (Alpha Channel)
            lngLength = UpperBounds(udtBitmap.ANDMap()) + 1
            lngChunkLength = lngChunkLength + lngLength
            Call WriteFile(hFile, udtBitmap.ANDMap(0), lngLength, lngBytesWritten, &H0)
            If Err.LastDllError Then
                Err.Raise menmErrorNumber.ERR_WRITING_FILE, udtErrInfo.Method, LastSystemError()
            End If
                            
            audtIconDOL(lngResCount).Length = lngChunkLength
            
            lngResCount = lngResCount + 1
        
        Next
    
    Next
    
    ' Pad the end of the file to double word boundary
    Call PadAmount(hFile, lngAlignShift, lngOffset)
        
    ' Now calculate the correct offsets and lengths
    For lngCounter = 0 To lngIconGroupCount - 1
        lngOffset = audtIconGroupDOL(lngCounter).offset
        audtIconGroupDOL(lngCounter).offset = lngOffset / lngAlignShift
        If ((lngOffset Mod lngAlignShift) > 0) Then
            audtIconGroupDOL(lngCounter).offset = audtIconGroupDOL(lngCounter).offset + 1
        End If
    Next
    For lngCounter = 0 To lngIconGroupCount - 2
        audtIconGroupDOL(lngCounter).Length = audtIconGroupDOL(lngCounter + 1).offset - audtIconGroupDOL(lngCounter).offset
    Next
    audtIconGroupDOL(lngIconGroupCount - 1).Length = (audtIconGroupDOL(lngIconGroupCount - 1).Length / lngAlignShift) + 1
         
    For lngCounter = 0 To lngIconCount - 1
        lngOffset = audtIconDOL(lngCounter).offset
        audtIconDOL(lngCounter).offset = lngOffset / lngAlignShift
        If ((lngOffset Mod lngAlignShift) > 0) Then
            audtIconDOL(lngCounter).offset = audtIconDOL(lngCounter).offset + 1
        End If
        
    Next
    For lngCounter = 0 To lngIconCount - 2
        audtIconDOL(lngCounter).Length = audtIconDOL(lngCounter + 1).offset - audtIconDOL(lngCounter).offset
    Next
    audtIconDOL(lngIconCount - 1).Length = (audtIconDOL(lngIconCount - 1).Length / lngAlignShift)
    
    ' --------------------------------------------------------------
    ' We need to go back and write out the Offset and Length data.
    udtErrInfo.AddlInfo = "Writing Icon Group TTypeInfo"
    
    Call SetFilePointer(hFile, lngIconGroupOffset, 0, FILE_BEGIN)
    
    udtTNameInfo.rnHandle = 0
    udtTNameInfo.rnUsage = 0
    udtTNameInfo.rnFlags = &H1C30
    lngLength = Len(udtTNameInfo) ' 12
    ReDim abytData(lngLength - 1)

    For lngCounter = 0 To lngIconGroupCount - 1
        ' Write out the UDT again to make sure the spacing is accurate.
        udtTNameInfo.rnOffset = CInt("&H" & Hex$(audtIconGroupDOL(lngCounter).offset)) ' UWORD
        udtTNameInfo.rnLength = CInt("&H" & Hex$(audtIconGroupDOL(lngCounter).Length)) ' UWORD
        udtTNameInfo.rnID = CInt("&H" & Hex$(RT_HIGHORDER + (lngCounter + 1)))
        
        ' Move the UDT to the byte array.
        Call CopyMemory(abytData(0), udtTNameInfo, lngLength)
        
        Call WriteFile(hFile, abytData(0), lngLength, lngBytesWritten, &H0)
        If Err.LastDllError Then
            Err.Raise menmErrorNumber.ERR_WRITING_FILE, udtErrInfo.Method, LastSystemError()
        End If
        
    Next
    
    udtErrInfo.AddlInfo = "Writing Icon TTypeInfo"
    Call SetFilePointer(hFile, lngIconOffset, 0, FILE_BEGIN)
        
    For lngCounter = 0 To lngIconCount - 1
        ' Write out the UDT again to make sure the spacing is accurate.
        udtTNameInfo.rnOffset = CInt("&H" & Hex$(audtIconDOL(lngCounter).offset)) ' UWORD
        udtTNameInfo.rnLength = CInt("&H" & Hex$(audtIconDOL(lngCounter).Length)) ' UWORD
        udtTNameInfo.rnID = CInt("&H" & Hex$(RT_HIGHORDER + (lngCounter + 1)))
        
        ' Move the UDT to the byte array.
        Call CopyMemory(abytData(0), udtTNameInfo, lngLength)
        
        Call WriteFile(hFile, abytData(0), lngLength, lngBytesWritten, &H0)
        If Err.LastDllError Then
            Err.Raise menmErrorNumber.ERR_WRITING_FILE, udtErrInfo.Method, LastSystemError()
        End If
        
    Next
        
    ' Write out the Align Shift value
    Call SetFilePointer(hFile, 192, 0, FILE_BEGIN)
    lngAlignShift = Log(lngAlignShift) / Log(2)
    Call WriteFile(hFile, lngAlignShift, 2, lngBytesWritten, &H0)
        
' <---------------------------------------------------------->
ExitCode:
    
    ' Flush the buffers and close the file
    If (hFile <> 0) Then
        Call FlushFileBuffers(hFile)
        Call CloseHandle(hFile)
    End If

    ' Deallocate all object references.
    Erase abytData()

    ' If an error occurs, raise the Err object and exit the function.
    If (udtErrInfo.Number <> menmErrorNumber.No_Error) Then
        Err.Raise udtErrInfo.Number, udtErrInfo.Source, udtErrInfo.Desc
    End If

    Exit Sub

' <---------------------------------------------------------->
ErrorCode:

    ' Set a permanent break point for the IDE. This has no effect in compiled code.
    Debug.Assert False

    ' Error occurred store it, stop error catching, and go to exit
    Call SetErrorInformation(Err.Number, Err.Source, Err.Description, Erl(), udtErrInfo)
    On Error GoTo 0
    Resume ExitCode
    Resume ' For Debug Only!

End Sub

' <---------------------------------------------------------------------------------------->
'
'     Module Name:  ICL.WriteByteToArray (Public)
'
'     Description:  Takes the passed-in byte array, and writes an unsigned byte to it.
'                   If the array is not large enough to hold the data at that offset, then
'                   expands it by 1 byte.
'
'      Dependancy:  (None)
'
'          Inputs:  vlngData,       ByVal Long, the data to write
'                   rabytData(),    ByRef Byte Array.
'                   rlngOffset,     ByRef Long, Offset to the current position in the array.
'
'         Outputs:  (None)
'
'  Change History:
'
'           Initials+#  Reference   Tracking #  Date        Description
'           ----------  ---------   ----------  ----------  ----------------------------------------------------
'           rda001      ICL         n/a         05/16/2003  Created for Project
'
Public Sub WriteByteToArray(ByVal vlngData As Long, ByRef rabytData() As Byte, ByRef rlngOffset As Long)

    On Error GoTo ErrorCode

Dim udtErrInfo  As mudtErrorData
Dim bytData     As Byte
Dim lngErr      As Long
Dim strData     As String

    udtErrInfo.Method = App.Title & mcstrMODULE_BASE_NAME & "WriteByteToArray()"
    
    ' Make the data unsigned
    strData = Right$(String$(2, "0") & Hex$(vlngData), 2)

    ' Make sure that data fits a byte data range
    bytData = CLng("&H" & strData) Mod 256

    On Error Resume Next
        rabytData(rlngOffset) = bytData
        lngErr = Err.Number
    On Error GoTo 0

    If (lngErr = 9) Then
        ' Subscript Out of Range
        ReDim Preserve rabytData(rlngOffset)
        rabytData(rlngOffset) = bytData
    End If

    rlngOffset = rlngOffset + 1

' <---------------------------------------------------------->
ExitCode:

    ' Deallocate all object references.

    ' If an error occurs, raise the Err object and exit the function.
    If (udtErrInfo.Number <> menmErrorNumber.No_Error) Then
        Err.Raise udtErrInfo.Number, udtErrInfo.Source, udtErrInfo.Desc
    End If

    Exit Sub

' <---------------------------------------------------------->
ErrorCode:

    ' Set a permanent break point for the IDE. This has no effect in compiled code.
    Debug.Assert False

    ' Error occurred store it, stop error catching, and go to exit
    Call SetErrorInformation(Err.Number, Err.Source, Err.Description, Erl(), udtErrInfo)
    On Error GoTo 0
    Resume ExitCode
    Resume ' For Debug Only!

End Sub

' <---------------------------------------------------------------------------------------->
'
'     Module Name:  ICL.CopyByteArray (Private)
'
'     Description:  Does the work of a Copy Memory of the Source array at element 0,
'                   to the Dest Array at element rlngOffset, for the length of the Source
'                   array, and sets rlngOffset equal to the new length of Dest + 1
'
'      Dependancy:  (None)
'
'          Inputs:  rabytDest,   ByRef Array of Bytes
'                   rabytSource, ByRef Array of Bytes
'                   rlngOffset,  ByRef Long
'
'         Outputs:  (None)
'
'  Change History:
'
'           Initials+#  Reference   Tracking #  Date        Description
'           ----------  ---------   ----------  ----------  ----------------------------------------------------
'           rda001      ICL         n/a         05/16/2003  Created for Project
'
Private Sub CopyByteArray(ByRef rabytDest() As Byte, ByRef rabytSource() As Byte, ByRef rlngOffset As Long)

    On Error GoTo ErrorCode

Dim udtErrInfo      As mudtErrorData
Dim lngLength       As Long
Dim lngNewLength    As Long

    udtErrInfo.Method = App.Title & mcstrMODULE_BASE_NAME & "CopyByteArray()"

    lngLength = UpperBounds(rabytSource()) + 1

    If (lngLength > 0) Then
        lngNewLength = lngLength + rlngOffset

        ReDim Preserve rabytDest(lngNewLength - 1)
        Call CopyMemory(rabytDest(rlngOffset), rabytSource(0), lngLength)
        rlngOffset = lngNewLength
    End If

' <---------------------------------------------------------->
ExitCode:

    ' Deallocate all object references.

    ' If an error occurs, raise the Err object and exit the function.
    If (udtErrInfo.Number <> menmErrorNumber.No_Error) Then
        Err.Raise udtErrInfo.Number, udtErrInfo.Source, udtErrInfo.Desc
    End If

    Exit Sub

' <---------------------------------------------------------->
ErrorCode:

    ' Error occurred store it, stop error catching, and go to exit
    Call SetErrorInformation(Err.Number, Err.Source, Err.Description, Erl(), udtErrInfo)
    On Error GoTo 0
    Resume ExitCode
    Resume ' For Debug Only!

End Sub

' <---------------------------------------------------------------------------------------->
'
'     Module Name:  ICL.UpperBounds (Private)
'
'     Description:  Returns the upper bounds on an array.
'
'      Dependency:  (None)
'
'          Inputs:  vvntArray, Variant holding an array
'                   ovlngDimension, Byval Optional Long, holds which array dimension to get the ubound
'                       on for multidimensional arrays. If = 0, then assumes a 1-d array.
'
'         Outputs:  -1 if the array is undefined, the upper bounds otherwise.
'
'  Change History:
'
'           Initials+#  Reference   Tracking #  Date        Description
'           ----------  ----------- ----------  ----------  ----------------------
'           rda001      ICL         n/a         05/16/2003  Created for Project
'
Private Function UpperBounds(ByVal vvntArray As Variant, Optional ByVal ovlngDimension As Long = 0) As Long

    On Error GoTo ErrorCode

Dim udtErrInfo  As mudtErrorData
Dim lngTotal    As Long

    udtErrInfo.Method = App.Title & mcstrMODULE_BASE_NAME & "UpperBounds()"

    If Not (IsArray(vvntArray)) Then
        Err.Raise menmErrorNumber.ERR_INVALID_PARAMETER, udtErrInfo.Method, "Invalid Array"
    End If

    On Error Resume Next
        lngTotal = -1
        If (ovlngDimension = 0) Then
            lngTotal = UBound(vvntArray)
        Else
            lngTotal = UBound(vvntArray, ovlngDimension)
        End If
    On Error GoTo ErrorCode

    UpperBounds = lngTotal

' <---------------------------------------------------------->
ExitCode:

    ' Deallocate all object references.

    ' If an error occurs, raise the Err object and exit the function.
    If (udtErrInfo.Number <> menmErrorNumber.No_Error) Then
        Err.Raise udtErrInfo.Number, udtErrInfo.Source, udtErrInfo.Desc, udtErrInfo.HelpFile, udtErrInfo.HelpContext
    End If

    Exit Function

' <---------------------------------------------------------->
ErrorCode:

    ' Set a permanent break point for the IDE. This has no effect in compiled code.
    Debug.Assert False
    
    ' Error occurred store it, stop error catching, and go to exit
    Call SetErrorInformation(Err.Number, Err.Source, Err.Description, Erl(), udtErrInfo)
    On Error GoTo 0
    Resume ExitCode
    Resume ' For Debug Only!

End Function

Private Sub PadAmount(ByVal hFile As Long, vlngAmount As Long, ByRef rlngOffset As Long)
    
Dim lngCounter As Long
Dim lngBytesWritten As Long

    For lngCounter = 1 To (vlngAmount - (rlngOffset Mod vlngAmount)) Mod vlngAmount
        Call WriteFile(hFile, 0&, 1, lngBytesWritten, &H0)
        rlngOffset = rlngOffset + 1
    Next

End Sub

' <---------------------------------------------------------------------------------------->
'
'     Module Name:  ICL.ReorderIconIDs (Private)
'
'     Description:  Goes to every Icon in each Icon group and sets its ID number.
'
'      Dependency:  (None)
'
'          Inputs:  (None)
'
'         Outputs:  (None)
'
'  Change History:
'
'           Initials+#  Reference   Tracking #  Date        Description
'           ----------  ----------- ----------  ----------  ----------------------
'           rda001      ICL         n/a         05/23/2003  Created for Project
'
Private Sub ReorderIconIDs()

    On Error GoTo ErrorCode

Dim udtErrInfo      As mudtErrorData
Dim lngCounter      As Long
Dim lngIconCounter  As Long
Dim lngID           As Long

    udtErrInfo.Method = App.Title & mcstrMODULE_BASE_NAME & "ReorderIconIDs()"

    lngID = 1
    For lngCounter = 0 To mlngIconCount - 1
        For lngIconCounter = 0 To maudtIconData(lngCounter).Group.DirHeader.idCount - 1
            maudtIconData(lngCounter).Group.Entries(lngIconCounter).nID = lngID
            lngID = lngID + 1
        Next
    Next

' <---------------------------------------------------------->
ExitCode:

    ' Deallocate all object references.

    ' If an error occurs, raise the Err object and exit the function.
    If (udtErrInfo.Number <> menmErrorNumber.No_Error) Then
        Err.Raise udtErrInfo.Number, udtErrInfo.Source, udtErrInfo.Desc, udtErrInfo.HelpFile, udtErrInfo.HelpContext
    End If

    Exit Sub

' <---------------------------------------------------------->
ErrorCode:

    ' Set a permanent break point for the IDE. This has no effect in compiled code.
    Debug.Assert False
    
    ' Error occurred store it, stop error catching, and go to exit
    Call SetErrorInformation(Err.Number, Err.Source, Err.Description, Erl(), udtErrInfo)
    On Error GoTo 0
    Resume ExitCode
    Resume ' For Debug Only!

End Sub

' <---------------------------------------------------------------------------------------->

Public Property Get IconPicture(ByVal lHDC As Long, ByVal vlngIconIndex As Long, ByVal vlngImageIndex As Long) As StdPicture

Dim hIcon As Long
   
   hIcon = IconHandle(lHDC, vlngIconIndex, vlngImageIndex)
   ' Change the Icon into a VB StdPicture object
   If (hIcon <> 0) Then
      Set IconPicture = IconToPicture(hIcon)
   End If
   
End Property

' <---------------------------------------------------------------------------------------->
'
'     Module Name:  ICL.IconToPicture (Private)
'
'     Description:  Returns a VB picture object containing the specified icon.
'                   Based off of code from Steve McMahon (steve@dogma.demon.co.uk)
'
'      Dependency:  (None)
'
'          Inputs:  hIcon, ByVal Long, Handle to the Icon
'
'         Outputs:  IPicture
'
'  Change History:
'
'           Initials+#  Reference   Tracking #  Date        Description
'           ----------  ----------- ----------  ----------  ----------------------
'           rda001      ICL         n/a         05/23/2003  Created for Project
'
Private Function IconToPicture(ByVal hIcon As Long) As IPicture
    
    On Error GoTo ErrorCode

Dim udtErrInfo  As mudtErrorData
Dim objNewPic   As Picture
Dim objPicConv  As PictDesc
Dim udtGUID     As Guid
    
    udtErrInfo.Method = App.Title & mcstrMODULE_BASE_NAME & "IconToPicture()"
    
    If hIcon = 0 Then
        GoTo ExitCode
    End If
        
    ' This is all magic if you ask me
    objPicConv.cbSizeofStruct = Len(objPicConv)
    objPicConv.picType = vbPicTypeIcon
    objPicConv.hImage = hIcon
    
    ' Fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    With udtGUID
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    Call OleCreatePictureIndirect(objPicConv, udtGUID, True, objNewPic)
    
    Set IconToPicture = objNewPic
    
' <---------------------------------------------------------->
ExitCode:

    ' Deallocate all object references.

    ' If an error occurs, raise the Err object and exit the function.
    If (udtErrInfo.Number <> menmErrorNumber.No_Error) Then
        Err.Raise udtErrInfo.Number, udtErrInfo.Source, udtErrInfo.Desc, udtErrInfo.HelpFile, udtErrInfo.HelpContext
    End If

    Exit Function

' <---------------------------------------------------------->
ErrorCode:

    ' Set a permanent break point for the IDE. This has no effect in compiled code.
    Debug.Assert False
    
    ' Error occurred store it, stop error catching, and go to exit
    Call SetErrorInformation(Err.Number, Err.Source, Err.Description, Erl(), udtErrInfo)
    On Error GoTo 0
    Resume ExitCode
    Resume ' For Debug Only!

End Function

' <---------------------------------------------------------------------------------------->
'
'     Module Name:  ICL.IconHandle (Private)
'
'     Description:  Returns an Icon Handle based on the Icon and Image indices passed in.
'                   Based off code found at vbAccelerator
'
'      Dependency:  (None)
'
'          Inputs:  lHDC, ByVal Long, Handle of the picture box to use.
'                   vlngIconIndex, ByVal Long
'                   vlngImageIndex, ByVal Long
'
'         Outputs:  Long
'
'  Change History:
'
'           Initials+#  Reference   Tracking #  Date        Description
'           ----------  ----------- ----------  ----------  ----------------------
'           rda001      ICL         n/a         05/23/2003  Created for Project
'
Private Function IconHandle(ByVal lHDC As Long, ByVal vlngIconIndex As Long, ByVal vlngImageIndex As Long) As Long

    On Error GoTo ErrorCode

Dim udtErrInfo  As mudtErrorData
Dim udtIconInfo As ICONINFO
    
    udtErrInfo.Method = App.Title & mcstrMODULE_BASE_NAME & "IconHandle()"
   
    ' Creates an icon indirectly from the bits
    udtIconInfo.fIcon = IMAGE_ICON
    udtIconInfo.hbmColor = GetIconBitmap(lHDC, vlngIconIndex, vlngImageIndex, ECFIImageConstants.ecfiImage, False, bReturnBmp:=True)
    udtIconInfo.hBmMask = GetIconBitmap(lHDC, vlngIconIndex, vlngImageIndex, ECFIImageConstants.ecfiMask, False, bReturnBmp:=True)
    
    IconHandle = CreateIconIndirect(udtIconInfo)
    
' <---------------------------------------------------------->
ExitCode:

    ' Deallocate all object references.
    Call DeleteObject(udtIconInfo.hbmColor)
    Call DeleteObject(udtIconInfo.hBmMask)

    ' If an error occurs, raise the Err object and exit the function.
    If (udtErrInfo.Number <> menmErrorNumber.No_Error) Then
        Err.Raise udtErrInfo.Number, udtErrInfo.Source, udtErrInfo.Desc, udtErrInfo.HelpFile, udtErrInfo.HelpContext
    End If

    Exit Function

' <---------------------------------------------------------->
ErrorCode:

    ' Set a permanent break point for the IDE. This has no effect in compiled code.
    Debug.Assert False
    
    ' Error occurred store it, stop error catching, and go to exit
    Call SetErrorInformation(Err.Number, Err.Source, Err.Description, Erl(), udtErrInfo)
    On Error GoTo 0
    Resume ExitCode
    Resume ' For Debug Only!

End Function

' <---------------------------------------------------------------------------------------->
'
'     Module Name:  ICL.GetIconBitmap (Private)
'
'     Description:  Based off code found at vbAccelerator
'
'      Dependency:  (None)
'
'          Inputs:  hIcon, ByVal Long, Handle to the Icon
'
'         Outputs:  Long
'
'  Change History:
'
'           Initials+#  Reference   Tracking #  Date        Description
'           ----------  ----------- ----------  ----------  ----------------------
'           rda001      ICL         n/a         05/28/2003  Created for Project
'           rda002      ICL         PSC 001     06/11/2003  Changed to use CreateBitmap thanks to advice
'                                                           from Carles (carles_pv@terra.es) at Planet-Source-Code.
'                                                           This will fix the Win98 masking problem.
'
Private Function GetIconBitmap(ByVal lHDC As Long, ByVal vlngIconIndex As Long, ByVal vlngImageIndex As Long, _
                               ByVal eType As ECFIImageConstants, Optional ByVal bDrawToDC As Boolean = True, _
                               Optional ByVal X As Long, Optional ByVal Y As Long, _
                               Optional ByVal lWidth As Long, Optional ByVal lHeight As Long, _
                               Optional ByVal eOp As VBRUN.RasterOpConstants, _
                               Optional ByVal bReturnBmp As Boolean = False) As Long
   
    On Error GoTo ErrorCode

Dim udtErrInfo  As mudtErrorData
Dim udtBMIMono  As BITMAPINFO_1BPP
Dim udtBMIXP    As BITMAPINFO_8BPP
Dim udtIconData As IconData
Dim udtMemEntry As MEMICONDIRENTRY
Dim udtBitmap   As IconBitmapData
Dim hDC         As Long
Dim lngRC       As Long
Dim hBmp        As Long
Dim hBmpOld     As Long
Dim abytData()  As Byte
Dim lngLength   As Long

    udtErrInfo.Method = App.Title & mcstrMODULE_BASE_NAME & "GetIconBitmap()"

    ' Make sure the Icon Index is a valid value.
    If (vlngIconIndex < 1) Or (vlngIconIndex > mlngIconCount) Then
        udtErrInfo.AddlInfo = CStr(vlngIconIndex)
        Err.Raise menmErrorNumber.ERR_DATA_OUT_OF_RANGE, udtErrInfo.Method, "Icon Index out of range"
    End If

    udtIconData = maudtIconData(vlngIconIndex - 1)
    
    ' Make sure the Image Index is a valid value.
    If (vlngImageIndex < 1) Or (vlngImageIndex > udtIconData.Group.DirHeader.idCount) Then
        udtErrInfo.AddlInfo = CStr(vlngImageIndex)
        Err.Raise menmErrorNumber.ERR_DATA_OUT_OF_RANGE, udtErrInfo.Method, "Image Index out of range"
    End If
    
    udtMemEntry = udtIconData.Group.Entries(vlngImageIndex - 1)
    udtBitmap = udtIconData.Bitmaps(vlngImageIndex - 1)
    
    ' Returns or draws a device dependent bitmap containing the Image (XOR)
    ' or Mask (AND) image from an icon
    If (eType = ECFIImageConstants.ecfiImage) Then
      
        ' Extract the XOR (colour) part of the icon image
   
        ' First create a compatible DC
        hDC = CreateCompatibleDC(lHDC) ' we can replace this with desktop DC
        If (hDC <> 0) Then
         
            ' Create a Bitmap compatible with the device
            hBmp = CreateCompatibleBitmap(lHDC, udtMemEntry.bWidth, udtMemEntry.bHeight)
            If (hBmp <> 0) Then
                ' Select the object into the DC
                hBmpOld = SelectObject(hDC, hBmp)
                If (hBmpOld <> 0) Then
                   
                    ' Icons have 2 x correct height, so temporarily correct the
                    ' BitmapInfoHeader structure whilst we create the bitmap
                    udtBitmap.Header.biHeight = udtBitmap.Header.biHeight \ 2
                   
                    ' Dump the bitmap information into a byte array to be used by SetDIBitsToDevice
                    
                    ' Check to see if there is a palette present
                    If (GetColorCount(udtBitmap.Header.biBitCount) > 0) Then
                        lngLength = ((UBound(udtBitmap.Palette()) + 1) * 4) + BITMAP_HEADER_SIZE
                        ReDim abytData(lngLength - 1)
                        Call CopyMemory(abytData(0), udtBitmap, lngLength)
                        Call CopyMemory(abytData(BITMAP_HEADER_SIZE), udtBitmap.Palette(0), lngLength - BITMAP_HEADER_SIZE)
                    Else
                        lngLength = BITMAP_HEADER_SIZE
                        ReDim abytData(lngLength - 1)
                        Call CopyMemory(abytData(0), udtBitmap, lngLength)
                    End If
                    
                    lngRC = SetDIBitsToDevice(hDC, 0, 0, udtMemEntry.bWidth, udtMemEntry.bHeight, _
                                              0, 0, 0, udtMemEntry.bHeight, udtBitmap.XORMap(0), _
                                              abytData(0), DIB_RGB_COLORS)
                    
                    ' Draw it if required
                    If (bDrawToDC) Then
                        If (lWidth = 0) And (lHeight = 0) Then
                            Call BitBlt(lHDC, X, Y, udtMemEntry.bWidth, udtMemEntry.bHeight, hDC, 0, 0, eOp)
                        Else
                            ' NB only allowing a larger version to be drawn here.
                            If (lWidth = 0) Then
                                lWidth = udtMemEntry.bWidth
                            End If
                            If (lHeight = 0) Then
                                lHeight = udtMemEntry.bHeight
                            End If
                            Call StretchBlt(lHDC, X, Y, lWidth, lHeight, hDC, 0, 0, udtMemEntry.bWidth, udtMemEntry.bHeight, eOp)
                        End If
                    End If
                   
                    ' Put the x2 icon height back again
                    udtBitmap.Header.biHeight = udtBitmap.Header.biHeight * 2
                    
                    ' Remove the bitmap from the DC
                    Call SelectObject(hDC, hBmpOld)
                End If
                
                ' Delete created bitmap if required
                If (bReturnBmp) Then
                   GetIconBitmap = hBmp
                Else
                   Call DeleteObject(hBmp)
                End If
            End If
             
            ' Clear up memory DC
            Call DeleteDC(hDC)
        End If
       
    ElseIf (ECFIImageConstants.ecfiMask) Then
      
        ' Extract the AND (mask) part of the icon image
      
        hDC = CreateCompatibleDC(0&)

        ' Create a monochrome DC
        hDC = CreateCompatibleDC(0&)
        If (hDC <> 0) Then
            
            ' Create a monochrome bitmap
            'hBmp = CreateCompatibleBitmap(lHDC, udtMemEntry.bWidth, udtMemEntry.bHeight)
            hBmp = CreateBitmap(udtMemEntry.bWidth, udtMemEntry.bHeight, 1, 1, ByVal 0&) ' rda002
            
            If (hBmp <> 0) Then
                ' Select the mono-bitmap into the DC
                hBmpOld = SelectObject(hDC, hBmp)
                If (hBmpOld <> 0) Then

                    ' We need to create a BitmapInfo structure that is a monochrome
                    ' version of the one provided in the Icon.
                    With udtBMIMono
                        With .bmiHeader
                            ' In icons the height = 2x the actual height
                            .biHeight = udtBitmap.Header.biHeight \ 2
                            .biWidth = udtBitmap.Header.biWidth
                            .biSize = udtBitmap.Header.biSize
                            .biPlanes = 1
                            .biBitCount = 1
                            .biCompression = BI_RGB
                            .biSizeImage = 0
                            .biXPelsPerMeter = 0
                            .biYPelsPerMeter = 0
                            .biClrUsed = 0
                            .biClrImportant = 0
                        End With
                        ' Set up monochrome colour palette
                        With .bmiColors(0)
                            .rgbRed = 0
                            .rgbGreen = 0
                            .rgbBlue = 0
                            .rgbReserved = 0
                        End With
                        With .bmiColors(1)
                            .rgbRed = 255
                            .rgbGreen = 255
                            .rgbBlue = 255
                            .rgbReserved = 0
                        End With
                    End With
                    
                    ' Set the Compatible Bitmap to the mask bits in the DIB within the icon
                    lngRC = SetDIBitsToDevice(hDC, 0, 0, udtMemEntry.bWidth, udtMemEntry.bHeight, _
                                              0, 0, 0, udtMemEntry.bHeight, udtBitmap.ANDMap(0), _
                                              udtBMIMono, DIB_RGB_COLORS)
                    
                    ' Draw it if required
                    If (bDrawToDC) Then
                        If (lWidth = 0) And (lHeight = 0) Then
                            Call BitBlt(lHDC, X, Y, udtMemEntry.bWidth, udtMemEntry.bHeight, hDC, 0, 0, eOp)
                        Else
                            ' NB only allowing a larger version to be drawn here.
                            If (lWidth = 0) Then
                                lWidth = udtMemEntry.bWidth
                            End If
                            If (lHeight = 0) Then
                                lHeight = udtMemEntry.bHeight
                            End If
                            Call StretchBlt(lHDC, X, Y, lWidth, lHeight, hDC, 0, 0, udtMemEntry.bWidth, udtMemEntry.bHeight, eOp)
                        End If
                    End If
                                 
                    ' Remove bitmap from DC
                    Call SelectObject(hDC, hBmpOld)
                  
                End If
            
                ' Clear up bitmap if required
                If (bReturnBmp) Then
                    GetIconBitmap = hBmp
                Else
                    Call DeleteObject(hBmp)
                End If
            End If
         
            ' Clear up memory DC.
            Call DeleteObject(hDC)
        End If
    End If
   
' <---------------------------------------------------------->
ExitCode:

    ' Deallocate all object references.
    Erase abytData()
    
    ' If an error occurs, raise the Err object and exit the function.
    If (udtErrInfo.Number <> menmErrorNumber.No_Error) Then
        Err.Raise udtErrInfo.Number, udtErrInfo.Source, udtErrInfo.Desc, udtErrInfo.HelpFile, udtErrInfo.HelpContext
    End If

    Exit Function

' <---------------------------------------------------------->
ErrorCode:

    ' Set a permanent break point for the IDE. This has no effect in compiled code.
    Debug.Assert False
    
    ' Error occurred store it, stop error catching, and go to exit
    Call SetErrorInformation(Err.Number, Err.Source, Err.Description, Erl(), udtErrInfo)
    On Error GoTo 0
    Resume ExitCode
    Resume ' For Debug Only!

End Function

Private Function GetColorCount(ByVal vlngBitCount As Long) As Long
    Select Case vlngBitCount
        Case 1:     GetColorCount = 2
        Case 4:     GetColorCount = 16
        Case 8:     GetColorCount = 256
        Case Else:  GetColorCount = 0
    End Select
End Function

Public Sub DrawIconImage(ByVal lHDC As Long, ByVal IconIndex As Long, ByVal ImageIndex As Long, _
                         Optional ByVal eType As Long = ecfiImage, _
                         Optional ByVal X As Long = 0, _
                         Optional ByVal Y As Long = 0, _
                         Optional ByVal lWidth As Long = 0, _
                         Optional ByVal lHeight As Long = 0, _
                         Optional ByVal eRasterOp As RasterOpConstants = vbSrcCopy)
   
    ' Draws either the Image (XOR) or Mask (AND) parts of the icon to an HDC:
    Call GetIconBitmap(lHDC, IconIndex, ImageIndex, eType, True, X, Y, lWidth, lHeight, eRasterOp)
    
End Sub


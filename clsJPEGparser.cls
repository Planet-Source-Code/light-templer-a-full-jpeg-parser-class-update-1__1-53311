VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJPEGparser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'
'   clsJPEGparser.cls
'

'   Started on:     7th April 2004
'   By        :     Light Templer
'   Update lvl:     1
'   Last edit :     22.04.2004



'   WHAT:
'   A selfcontained class done from scratch to parse the segments of
'   a JPEG file and get some informations (like the TRUE size (width /
'   height) of the image when there is a Adobe PhotoShop (TM) preview
'   included ..., start of segments, parseable without errors and so on)


'   CREDITS:
'
'   Many thx to Christian Tratz for sharing his JPEG informations
'   on www.codeproject.com in his Visual-C project 'Extracting
'   IPTC header information from JPEG images' !
'
'   Eric Hamilton for his JPEG structure docu from 1992.
'
'   Some more JPEG informations used taken from several web sites.
'
'   API information mostly taken from www.ALLAPI.NET
'   Thx for their free fine app/lib  API-Guide / API-Viewer !


'   COPYRIGHT / CONTACT
'
'   All (C) by Light Templer. Please send any problems / improvements
'   to   schwepps_bitterlemon@gmx.de


'   ===========================================
'   EXAMPLE:
'
'   Dim oJPEGparser as clsJPEGparser
'
'   Set oJPEGparser = new clsJPEGparser
'
'   oJPEGparser.ParseJpegFile("PathToYourJPEGFile")
'   Debug.Print oJPEGparser.FullReport
'
'   '   Debug.Print oJPEGparser. ... "many other, try it!"
'
'   Set oJPEGparser = Nothing
'   ===========================================


'   UPDATE 1:
'               - Added missing properties for colordepth (as value and as text)
'               - Added MARKER_SOF1 to MARKER_SOF7 as Start-Of-Frame marker to get the size and colordepth.
'                 So we handle progressive encoded JPEGS correctly now.


Option Explicit



' *************************************
' *            CONSTANTS              *
' *************************************
Private Const MIN_SIZE_JPEG_FILE = 250                  ' Used for error checking. AFAIK a JPEG file
                                                        ' cannot be smaller than this. (Checked using
                                                        ' QuickView with a one pixel sized gray image,
                                                        ' change if you want ;) !)
                                                        
' Markers in a JPEG file for segments with "standard" structure:  Marker, Size , Data
Private Const MARKER_SOI = "FFD8"           ' Start-Of-Image
Private Const MARKER_EOI = "FFD9"           ' End-Of-Image

Private Const MARKER_APP0 = "FFE0"          ' Application marker 0 (there are APP0 to APP15 - FFE0 to FFEF)
Private Const MARKER_EXIF = "FFE1"          ' Exif header (mostly written by digital cameras, scanners, ...)

Private Const MARKER_APP2 = "FFE2"          ' Used by ???
Private Const MARKER_APP3 = "FFE3"          ' Used by ???
Private Const MARKER_APP4 = "FFE4"          ' Used by ???
Private Const MARKER_APP5 = "FFE5"          ' Used by ???
Private Const MARKER_APP6 = "FFE6"          ' Used by ???
Private Const MARKER_APP7 = "FFE7"          ' Used by ???
Private Const MARKER_APP8 = "FFE8"          ' Used by ???
Private Const MARKER_APP9 = "FFE9"          ' Used by ???
Private Const MARKER_APP10 = "FFEA"         ' Used by ???
Private Const MARKER_APP11 = "FFEB"         ' Used by ???
Private Const MARKER_APP12 = "FFEC"         ' Used by ???

Private Const MARKER_APP14 = "FFED"         ' APP14 (Adobe PhotoShop (TM) picture informations) ---  btw: Don't know
                                            ' why its called APP14; imho should be APP13 ...
                                            
Private Const MARKER_APPEE = "FFEE"         ' APP?? (Seen in JPEGs written by Adobe PhotoShop (TM))
Private Const MARKER_APP15 = "FFEF"         ' APP15

Private Const MARKER_DQT = "FFDB"           ' Quantization table
Private Const MARKER_DHT = "FFC4"           ' Huffman table
Private Const MARKER_SOF0 = "FFC0"          ' Start of frame
Private Const MARKER_SOS = "FFDA"           ' Start of scan
Private Const MARKER_DRI = "FFDD"           ' Define restart interval
Private Const MARKER_COM = "FFFE"           ' Comments

' Type of coding (Baseline/Progressiv, ...) markers in a JPEG file. Segments with "standard" structure:  Marker, Size , Data
Private Const MARKER_SOF1 = "FFC1"          ' Extended sequential DCT, Huffman
Private Const MARKER_SOF2 = "FFC2"          ' Progressive DCT, Huffman
Private Const MARKER_SOF3 = "FFC3"          ' Spatial (sequential) lossless, Huffman
Private Const MARKER_SOF5 = "FFC5"          ' Differential Sequential DCT, Huffman
Private Const MARKER_SOF6 = "FFC6"          ' Differential progressive DCT, Huffman
Private Const MARKER_SOF7 = "FFC7"          ' Differential spatial, Huffman
Private Const MARKER_SOF9 = "FFC9"          ' Extended sequential DCT, Arithmetic
Private Const MARKER_SOF10 = "FFCA"         ' Progressive DCT, Arithmetic
Private Const MARKER_SOF11 = "FFCB"         ' Spatial (sequential) lossless, Arithmetic
Private Const MARKER_SOF13 = "FFCD"         ' Differential sequential DCT, Arithmetic
Private Const MARKER_SOF14 = "FFCE"         ' Differential progressive DCT, Arithmetic
Private Const MARKER_SOF15 = "FFCF"         ' Differential spatial, Arithmetic

Private Const MARKER_JPG = "FFC8"
Private Const MARKER_DAC = "FFCC"           ' Define Arithmetic coding conditioning
Private Const MARKER_DNL = "FFDC"           ' Define number of Lines
Private Const MARKER_DHP = "FFDE"           ' Define Hierarchical progression
Private Const MARKER_EXP = "FFDF"           ' Expand reference components
Private Const MARKER_JPG0 = "FFF0"          ' Reserved for JPEG extensions
'                    ... to ...
Private Const MARKER_JPG13 = "FFFD"         ' Reserved for JPEG extensions

' Markers in a JPEG file for segments without any further data, just the two marker bytes
Private Const MARKER_TEM = "FF01"           ' "Usually causes a decoding error, may be ignored"
Private Const MARKER_RTS0 = "FFD0"          ' RSTn are used for resync. You can find them within
Private Const MARKER_RTS1 = "FFD1"          ' a SOS segment only!
Private Const MARKER_RTS2 = "FFD2"
Private Const MARKER_RTS3 = "FFD3"
Private Const MARKER_RTS4 = "FFD4"
Private Const MARKER_RTS5 = "FFD5"
Private Const MARKER_RTS6 = "FFD6"
Private Const MARKER_RTS7 = "FFD7"



' *************************************
' *        API DEFINITIONS            *
' *************************************
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const MAX_PATH As Long = 260

Private Type FILETIME
    dwLowDateTime       As Long
    dwHighDateTime      As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes    As Long
    ftCreationTime      As FILETIME
    ftLastAccessTime    As FILETIME
    ftLastWriteTime     As FILETIME
    nFileSizeHigh       As Long
    nFileSizeLow        As Long
    dwReserved0         As Long
    dwReserved1         As Long
    cFileName           As String * MAX_PATH
    cAlternate          As String * 14
End Type

Private Declare Function API_FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
        (ByVal lpFileName As String, _
         lpFindFileData As WIN32_FIND_DATA) As Long
   
Private Declare Function API_FindClose Lib "kernel32" Alias "FindClose" _
        (ByVal hFindFile As Long) As Long


Private Declare Function API_CreateFile Lib "kernel32" Alias "CreateFileA" _
        (ByVal lpFileName As String, _
         ByVal dwDesiredAccess As Long, _
         ByVal dwShareMode As Long, _
         ByVal lpSecurityAttributes As Long, _
         ByVal dwCreationDisposition As Long, _
         ByVal dwFlagsAndAttributes As Long, _
         ByVal hTemplateFile As Long) As Long

Private Declare Function API_GetFileSize Lib "kernel32" Alias "GetFileSize" _
        (ByVal hFile As Long, _
         lpFileSizeHigh As Long) As Long

Private Declare Function API_ReadFile Lib "kernel32" Alias "ReadFile" _
        (ByVal hFile As Long, _
         lpBuffer As Any, _
         ByVal nNumberOfBytesToRead As Long, _
         lpNumberOfBytesRead As Long, _
         ByVal lpOverlapped As Any) As Long

Private Declare Function API_CloseHandle Lib "kernel32" Alias "CloseHandle" _
        (ByVal hObject As Long) As Long




' *************************************
' *            PRIVATES               *
' *************************************

Private arrByteJPEGfile()   As Byte         ' Here we hold the whole file in memory to parse it.
                                            ' Size should be no problem - JPEGs are small.
                                            
                                            
Private Type tpMvar                         ' UDT to hold all local informations in one easy to access var
                                                
    IndexApp0Header         As Long
    IndexEXIFHeader         As Long
    IndexPhotoShopComments  As Long         ' IPTC comments
    IndexPictureHeader      As Long
    IndexPictureimage       As Long
        
    HasPreview              As Boolean
    HasEXIFHeader           As Boolean
    HasPhotoShopComments    As Boolean
    ParsedWithoutProblems   As Boolean
    
    Filename                As String
    Filesize                As Long
    JPEGVersion             As String
    HowManyComments         As Long
    Comments()              As String       ' 1-based array with (1 to HowManyComments) elements
    FullReport              As String
    Resolution              As String
    
    XsizePreview            As Long
    YsizePreview            As Long
    XsizePicture            As Long
    YsizePicture            As Long
    ColorDepthInBit         As Long
    ColorDepthAsText        As String
    
    ErrorMsg                As String
    
End Type
Private mvar As tpMvar                      ' Holds all local informations in a handy way.
'
'
'



' *************************************
' *            INIT/TERM              *
' *************************************

Private Sub Class_Initialize()

    ' nothing yet ...

End Sub

Private Sub Class_Terminate()

    ' nothing yet ...

End Sub




' *************************************
' *         PUBLIC FUNCTIONS          *
' *************************************

Public Function ParseJpegFile(PathFilename As String) As Boolean
    ' Returns TRUE, when loaded successfully and parsed.
    
    ResetAll
    
    With mvar
        .Filename = PathFilename
        
        ' CHECK:  Does file exists and is it a valid file (not a directory ...)
        If DoesFileExists(PathFilename) = False Then
            .ErrorMsg = "Not a valid file:  '" + PathFilename + "' !"
            
            Exit Function
        End If
                
        ' READ:  the whole file into our local buffer byte array  'arrByteJPEGfile()'
        If ReadFileIntoBuffer(PathFilename) = False Then
        
            Exit Function
        End If
            
        ' PARSE:  Here we do the job
        If ParseFile() = False Then
        
            Exit Function
        End If
                
    End With
    
    ParseJpegFile = True
    
End Function

Public Sub Clear()
    ' Clears byte array with JPEG file and any information about it!
        
    ResetAll
       
End Sub




' *************************************
' *         PRIVATE FUNCTIONS         *
' *************************************

Private Function ResetAll()

    Dim EmptyDummyType As tpMvar

    Let mvar = EmptyDummyType       ' Quick erase a type var ;)
    Erase arrByteJPEGfile()
    
End Function


Private Function DoesFileExists(sPathFilename As String) As Boolean
    
    Dim hFile   As Long
    Dim WFD     As WIN32_FIND_DATA
    
    
    hFile = API_FindFirstFile(sPathFilename, WFD)
    If hFile <> INVALID_HANDLE_VALUE Then
        DoesFileExists = IIf((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY, _
                False, True)
    End If
    Call API_FindClose(hFile)

End Function


Private Function ReadFileIntoBuffer(sPathFilename As String) As Boolean
    ' ALL done with API calls to avoid VB errors (sensefull for NT services ...)
    
    Const OPEN_EXISTING = 3
    Const FILE_SHARE_READ = &H1
    Const GENERIC_READ = &H80000000
    Const FILE_BEGIN = 0
    Const READ_ERROR = 0
    
    Dim hFile   As Long
    Dim lBytes  As Long
    
    
    On Local Error GoTo error_handler

    With mvar
        
        ' OPEN:   the file ...
        hFile = API_CreateFile(sPathFilename, _
                GENERIC_READ, _
                FILE_SHARE_READ, _
                ByVal 0&, _
                OPEN_EXISTING, _
                0, 0)
    
        If hFile = INVALID_HANDLE_VALUE Then                ' ... leave on error! ...
            
            Exit Function
        End If
        
        .Filesize = API_GetFileSize(hFile, 0)               ' and get the filesize.
        
        ' CHECK:  Large enough to be a valid JPEG?
        If .Filesize < MIN_SIZE_JPEG_FILE Then
            .ErrorMsg = "To small to be a valid JPEG file:  '" + sPathFilename + "' !"
            API_CloseHandle hFile
            
            Exit Function
        End If
        
        ' READ: Resize the buffer and read the whole file into it
        ReDim arrByteJPEGfile(1 To .Filesize) As Byte
        If API_ReadFile(hFile, arrByteJPEGfile(1), .Filesize, lBytes, ByVal 0&) = READ_ERROR Then
            API_CloseHandle hFile
            .ErrorMsg = "Read error on:  '" + sPathFilename + "' !"
            
            Exit Function
        End If
    
        If lBytes <> .Filesize Then
            API_CloseHandle hFile
            .ErrorMsg = "Want to read " & .Filesize & " bytes from '" + sPathFilename + "', " & _
                    "got only " & lBytes & " !"
            
            Exit Function
        End If
        
        ' CLOSE: Successfully read the whole file, now close it.
        API_CloseHandle hFile
    
    End With
       
    ReadFileIntoBuffer = True

    Exit Function



error_handler:

    mvar.ErrorMsg = "[" + Err.Description + "] in procedure 'ReadFileIntoBuffer()' at class clsJPEGparser"

End Function


Private Function ParseFile() As Boolean
    ' MAIN LOOP:  Here all parsing of the filled buffer is done
    
    Dim lIndex          As Long         ' Pointer into buffer array:  Current position
    Dim lTmpIndex       As Long         ' Pointer into buffer array:  Current position in a segment
    Dim lPtrEnd         As Long         ' Pointer into buffer array:  position of last integer value (array size -1)
    Dim sSegmentMarker  As String       ' 4 chars in hex e.g. 'FFDB' is the marker for the quantization table
    Dim lSegmentSize    As Long         ' Counted WITHOUT the two marker bytes!
    Dim bytChar         As Byte         ' A single byte multiple used
    Dim sString         As String       ' Any kind of text - multiple used
    
    On Local Error GoTo error_handler

    
    lPtrEnd = UBound(arrByteJPEGfile()) - 1
    
    With mvar
            
        AddToFullReport "Filename: " & .Filename                                    ' Hint:  All adding to the report is
        AddToFullReport "Filesize: " & Format(.Filesize, "#,#") & " bytes" & vbCrLf ' done in this function or "below" it
        
        
        ' CHECK: Is it a valid JPEG file ?   (Must start with 'FF D8' and end with 'FF D9' )
        If GetIntegerAsHex(1) <> MARKER_SOI Or GetIntegerAsHex(lPtrEnd) <> MARKER_EOI Then
            .ErrorMsg = "JPEG start (SOI-FFD8) or end markers (EOI-FFD9) not valid !"
            
            Exit Function
        End If
        AddToFullReport "Start-Of-Image and End-Of-Image markers are ok" & vbCrLf
        
        
        ' BIG LOOP:  Jump from segment to segment, scan type and handle (some of) them
        lIndex = 3  ' Start of first segment
        
        Do While lIndex < lPtrEnd
        
            ' READ:  Header (Marker and size) of current segment
            sSegmentMarker = GetIntegerAsHex(lIndex)
            lSegmentSize = GetIntegerValue(lIndex + 2)
            
            
            ' SKIP:  Any number of FF bytes between two segments are legal. We must skip them here.
            If sSegmentMarker = "FFFF" Then
                Do While arrByteJPEGfile(lIndex + 1) = &HFF
                    lIndex = lIndex + 1
                Loop
                lSegmentSize = GetIntegerValue(lIndex + 2)
            End If
            
            Select Case sSegmentMarker
                            
                Case MARKER_APP0        ' Application marker
                        .IndexApp0Header = lIndex
                        
                        ' Get JPEG version
                        .JPEGVersion = arrByteJPEGfile(lIndex + 9) & "." & Right$("00" & arrByteJPEGfile(lIndex + 10), 2)
                        
                        ' Get resolution or aspect ratio
                        bytChar = arrByteJPEGfile(lIndex + 11)
                        If bytChar = 0 Then     ' aspect ratio
                            .Resolution = "Aspect ratio width/height = " & GetIntegerValue(lIndex + 12) & " : " & _
                                    GetIntegerValue(lIndex + 14)
                        
                        ElseIf bytChar = 1 Then
                            .Resolution = GetIntegerValue(lIndex + 12) & " x " & _
                                    GetIntegerValue(lIndex + 14) & " Dots/inch (DPI)"
                        
                        ElseIf bytChar = 2 Then
                            .Resolution = GetIntegerValue(lIndex + 12) & " x " & _
                                    GetIntegerValue(lIndex + 14) & " Dots/cm"
                                        
                        Else
                            .Resolution = "Error resolving resolution. Opcode is " & bytChar & _
                                    "Valide opcodes are 0, 1, 2"
                            
                        End If
                        
                        ' Check for standard JPEG thumbnails (they are rare, but anyway... )
                        If GetIntegerValue(lIndex + 16) <> 0 Then
                            .HasPreview = True
                            .XsizePreview = arrByteJPEGfile(lIndex + 16)
                            .YsizePreview = arrByteJPEGfile(lIndex + 17)
                        End If
                        
                        AddToFullReport MARKER_APP0 & " - Application marker (APP0)" & _
                                FormatPositionAndLength(lIndex, lSegmentSize) & vbCrLf & _
                                "    JPEG version: " & .JPEGVersion & vbCrLf & _
                                "    Resolution: " & .Resolution & vbCrLf & _
                                IIf(.HasPreview = True, _
                                        "    Preview width: " & .XsizePreview & " pixels" & vbCrLf & _
                                        "    Preview height: " & .YsizePreview & " pixels" & vbCrLf, "")
                        
                        
                                                     
                Case MARKER_APP2, MARKER_APP3, MARKER_APP4, MARKER_APP5, MARKER_APP6, MARKER_APP7, MARKER_APP8, _
                        MARKER_APP9, MARKER_APP10, MARKER_APP11, MARKER_APP12, MARKER_APP15
                        
                        ' Those are applikation markers some companies maybe use (like Adobe (TM) )
                        
                        AddToFullReport sSegmentMarker & " - APP" & Format("&H" & Right$(sSegmentMarker, 1)) & _
                                " segment" & FormatPositionAndLength(lIndex, lSegmentSize) & vbCrLf
                                                                                                
                                                                                                
                        
                Case MARKER_APP14       ' Adobe PhotoShop (TM) picture informations
                
                        ' HINT 1: Parsing an PhotoShop header isn't an easy task! AFAIK there's only a C++ source for.
                        
                        ' Hint 2: Beside the JPEG standard thumbnails Adobe specified an own version. Beside many
                        '         picture describing comments they put a whole FULL JPEG picture into this segment.
                        '         Yes, with FFD8 at beginning and FFD9 at the end. Nested, nested, ...
                        
                        .HasPhotoShopComments = True
                        .IndexPhotoShopComments = lIndex
                        
                        AddToFullReport MARKER_APP14 & " - Adobe PhotoShop (TM) picture informations" & _
                                FormatPositionAndLength(lIndex, lSegmentSize) & vbCrLf
                                                
                        
                        
                Case MARKER_APPEE       ' not an "official" name! Creation by me ;)
                        AddToFullReport MARKER_APPEE & " - #Unknown app segment - seen in JPEGs written by Adobe PhotoShop (TM)" & _
                                 FormatPositionAndLength(lIndex, lSegmentSize) & vbCrLf
                                                                
                                
                                
                Case MARKER_EXIF        ' Exif header
                        .HasEXIFHeader = True
                        .IndexEXIFHeader = lIndex
                        
                        ' HINT: Parsing an EXIF header isn't an easy task! But there's a VB class for (ExifReader).
                        
                        AddToFullReport MARKER_EXIF & " - Exif header" & _
                                FormatPositionAndLength(lIndex, lSegmentSize) & vbCrLf
                                
                
                
                Case MARKER_DRI         ' Define restart interval (always 4 bytes length)
                        AddToFullReport MARKER_DRI & " - Restart interval marker" & _
                                FormatPositionAndLength(lIndex, lSegmentSize) & " (always 4 bytes)" & vbCrLf
                
                
                
                Case MARKER_COM         ' Comments
                        sString = GetStringZeroTerm(lIndex + 4)
                        If sString <> "" Then
                            .HowManyComments = .HowManyComments + 1
                            ReDim Preserve .Comments(1 To .HowManyComments)
                            .Comments(.HowManyComments) = sString
                        End If
                        
                        AddToFullReport MARKER_COM & " - Comment header" & _
                                FormatPositionAndLength(lIndex, lSegmentSize) & vbCrLf & _
                                "    Comment: '" + sString + "'" + vbCrLf
                        
                        
                        
                Case MARKER_SOF0, MARKER_SOF1, MARKER_SOF2, MARKER_SOF3, _
                        MARKER_SOF5, MARKER_SOF6, MARKER_SOF7               ' Start of frame (the image "header")
                                                                            ' mabe here we need MARKER_SOF7 to
                                                                            ' MARKER_SOF15, too. Don' know so far!
                
                        .IndexPictureHeader = lIndex
                        
                        ' Get pics dimensions
                        .YsizePicture = GetIntegerValue(lIndex + 5)
                        .XsizePicture = GetIntegerValue(lIndex + 7)
                        
                        ' Get colore depth as values and text
                        .ColorDepthInBit = arrByteJPEGfile(lIndex + 9) * 8
                        If .ColorDepthInBit = 8 Then
                            .ColorDepthAsText = "Grayscale 8 Bit"
                            
                        ElseIf .ColorDepthInBit = 24 Then
                            .ColorDepthAsText = "RGB 24 Bit"
                            
                        ElseIf .ColorDepthInBit = 32 Then
                            .ColorDepthAsText = "CMYK 32 Bit"
                        
                        Else
                            .ColorDepthAsText = "Unknown format with " & .ColorDepthInBit & " Bit"
                            
                        End If
                        
                        AddToFullReport sSegmentMarker & " - Start of frame" & _
                                FormatPositionAndLength(lIndex, lSegmentSize) & vbCrLf & _
                                "  Length: " & Format(lSegmentSize, "#,#") & " Bytes" & vbCrLf & _
                                "  Pic width: " & .XsizePicture & " pixels" & vbCrLf & _
                                "  Pic height: " & .YsizePicture & " pixels" & vbCrLf & _
                                "  Color depth: " & .ColorDepthAsText & vbCrLf
                                
                        
                
                Case MARKER_SOS         ' Start of scan (with the image data)
                        ' Contains header AND image data, here the size shows end of header only!
                        ' We have to skip the image data to get the next segment... !
                        
                        .IndexPictureimage = lIndex
                        
                        ' HINT1:  FF 00 means FF . The following 00 is a flag for a "true" FF.
                        ' HINT2:  We have to skip the resync markers (RTS0 to RTS7), too!
                        
                        ' Skip the header
                        lTmpIndex = lIndex + lSegmentSize
                        
                        ' Walking byte for byte through the image data to find next marker
                        Do
                            If arrByteJPEGfile(lTmpIndex) = &HFF Then
                                bytChar = arrByteJPEGfile(lTmpIndex + 1)
                                
                                If bytChar >= &HD0 And bytChar <= &HD7 Then
                                    ' Skip the resync marker
                                    lTmpIndex = lTmpIndex + 1
                                    
                                ElseIf bytChar = &H0 Then
                                    ' Skip the "true" FF
                                    lTmpIndex = lTmpIndex + 1
                                    
                                Else
                                    ' Not a resync and not a "true" FF? We reached the end and leave the loop!
                                    lSegmentSize = lTmpIndex - (lIndex + 2)
                                    
                                    Exit Do
                                End If
                            End If
                            
                            lTmpIndex = lTmpIndex + 1
                        Loop
                        
                        AddToFullReport MARKER_SOS & " - Start of scan" & _
                                FormatPositionAndLength(lIndex, lSegmentSize) & vbCrLf
                                                        
                        
                        
                        
                Case MARKER_DQT         ' Quantization table (There can be one or more!)
                        AddToFullReport MARKER_DQT & " - Quantization table" & _
                                FormatPositionAndLength(lIndex, lSegmentSize) & vbCrLf
                        
                        
                        
                        
                Case MARKER_DHT         ' Huffman table (There can be one or more!)
                        AddToFullReport MARKER_DHT & " - Huffman table" & _
                                FormatPositionAndLength(lIndex, lSegmentSize) & vbCrLf
                
                
                
                
                Case MARKER_SOF2, MARKER_SOF3, MARKER_SOF5, MARKER_SOF6, MARKER_SOF7, MARKER_SOF9, _
                        MARKER_SOF10, MARKER_SOF11, MARKER_SOF13, MARKER_SOF14, MARKER_SOF15, _
                        MARKER_JPG, MARKER_DAC, MARKER_DNL, MARKER_DHP, MARKER_EXP, MARKER_JPG0, _
                        MARKER_JPG13    ' Usually unsupported markers in a JPEG file for segments
                                        ' with "standard" structure:  Marker, Size , Data
                
                        AddToFullReport sSegmentMarker & " - #Usually unsupported marker#" & _
                                 FormatPositionAndLength(lIndex, lSegmentSize) & vbCrLf
                                 
                
                
                Case Else               ' Any other not yet implemented segements
                        AddToFullReport sSegmentMarker & " - #Unknown segment type#" & _
                                 FormatPositionAndLength(lIndex, lSegmentSize) & vbCrLf
                        
            
            End Select
            
            ' JUMP: To start of next segment
            lIndex = lIndex + lSegmentSize + 2
        Loop
        
        If GetIntegerAsHex(lIndex) = MARKER_EOI Then
            ' Yes, we have reached the end and it is the JPEG End-Of-File marker, so wen can say:
            .ParsedWithoutProblems = True
            AddToFullReport "--> File successfully parsed." & vbCrLf
        End If
    End With
            
    ParseFile = True

    Exit Function


error_handler:
    
    
    mvar.ErrorMsg = "[" + Err.Description + "] in procedure 'ParseFile()' at class 'clsJPEGparser'"

End Function

Private Function GetIntegerValue(lIndex As Long) As Long
    ' Reads from position 'lIndex' two bytes and returns them as value (HighValue * 256 + LowValue)

    GetIntegerValue = CLng(arrByteJPEGfile(lIndex)) * 256& + CLng(arrByteJPEGfile(lIndex + 1))

End Function

Private Function GetIntegerAsHex(lIndex As Long) As String
    ' Reads from position 'lIndex' two bytes and returns them as a 4-digit hex string, e.g. "FFD9"

    GetIntegerAsHex = Right$("00" & Hex$(arrByteJPEGfile(lIndex)), 2) & _
            Right$("00" & Hex$(arrByteJPEGfile(lIndex + 1)), 2)

End Function

Private Function GetStringZeroTerm(ByRef Return_lIndex As Long) As String
    ' Reads from position 'Return_lIndex' a zero terminated string
    ' After reading Return_lIndex points to the byte AFTER the terminating zero byte!
           
    Dim lEndOfArr As Long
    
    lEndOfArr = UBound(arrByteJPEGfile())
    Do While arrByteJPEGfile(Return_lIndex) <> 0
        
        If Return_lIndex >= lEndOfArr Then
            mvar.ErrorMsg = "String doesn't end at file end!"       ' maybe here an event handler would be better...
            
            Exit Function
        End If
        
        GetStringZeroTerm = GetStringZeroTerm + Chr$(arrByteJPEGfile(Return_lIndex))
            
        Return_lIndex = Return_lIndex + 1
    Loop
    
    Return_lIndex = Return_lIndex + 1
    
End Function

Private Function FormatPositionAndLength(lIndex As Long, lLength As Long) As String
    ' Here we have ONE central position to style the output format of the 'fullreport' - Change to your needs!
    
    FormatPositionAndLength = " starts at: 0x" & Hex$(lIndex) & " / " & Format(lIndex, "#,#") & _
            "  Seg.length: 0x" & Hex$(lLength) & " / " & Format(lLength, "#,#") & " bytes"

End Function

Private Function AddToFullReport(sText As String)
    ' Here we build a "printable" string with all informations we get during parsing the segments
    
    mvar.FullReport = mvar.FullReport + sText + vbCrLf

End Function





' *************************************
' *           PROPERTIES              *
' *************************************


' Hint: All indices are 1-based (a file starts with the first (One) byte for me ;) )

Public Property Get IndexApp0Header() As Long
    
    IndexApp0Header = mvar.IndexApp0Header
 
End Property


Public Property Get IndexPhotoShopComments() As Long
    
    IndexPhotoShopComments = mvar.IndexPhotoShopComments
 
End Property

Public Property Get IndexPictureHeader() As Long
    
    IndexPictureHeader = mvar.IndexPictureHeader
 
End Property

Public Property Get IndexPictureimage() As Long
    
    IndexPictureimage = mvar.IndexPictureimage
 
End Property

Public Property Get IndexEXIFHeader() As Long
    
    IndexEXIFHeader = mvar.IndexEXIFHeader
 
End Property

Public Property Get HasPhotoShopComments() As Boolean
    
    HasPhotoShopComments = mvar.HasPhotoShopComments
 
End Property

Public Property Get HasEXIFHeader() As Boolean
    
    HasEXIFHeader = mvar.HasEXIFHeader
 
End Property

Public Property Get HasPreview() As Boolean
        
    ' HINT: Right now we can detect JPEG standard thumbnails (previews) only, NOT Adobe PhotoShop (TM) ones ...
    
    HasPreview = mvar.HasPreview
 
End Property

Public Property Get ParsedWithoutProblems() As Boolean
    
    ParsedWithoutProblems = mvar.ParsedWithoutProblems
 
End Property

Public Property Get Filename() As String
    
    Filename = mvar.Filename
 
End Property

Public Property Get Filesize() As Long
    
    Filesize = mvar.Filesize
 
End Property

Public Property Get JPEGVersion() As String
    
    JPEGVersion = mvar.JPEGVersion
 
End Property

Public Property Get HowManyComments() As Long
    
    HowManyComments = mvar.HowManyComments
 
End Property

Public Property Get Comments(lIndex As Long) As String
    
    With mvar
        If .HowManyComments > 0 Then
            If lIndex > 0 And lIndex <= UBound(.Comments()) Then
                Comments = .Comments(lIndex)
            End If
        End If
    End With
 
End Property

Public Property Get Resolution() As String
    
    Resolution = mvar.Resolution
 
End Property

Public Property Get FullReport() As String
    
    FullReport = mvar.FullReport
 
End Property

Public Property Get XsizePicture() As Long
    
    XsizePicture = mvar.XsizePicture
 
End Property

Public Property Get YsizePicture() As Long
    
    YsizePicture = mvar.YsizePicture
 
End Property

Public Property Get XsizePreview() As Long
    
    XsizePreview = mvar.XsizePreview
 
End Property

Public Property Get YsizePreview() As Long
    
    YsizePreview = mvar.YsizePreview
 
End Property

Public Property Get ColorDepthInBit() As Long
    
    ColorDepthInBit = mvar.ColorDepthInBit
 
End Property

Public Property Get ColorDepthAsText() As String
    
    ColorDepthAsText = mvar.ColorDepthAsText
 
End Property




Public Property Get ErrorMsg() As String
    
    ErrorMsg = mvar.ErrorMsg
 
End Property


' #*#

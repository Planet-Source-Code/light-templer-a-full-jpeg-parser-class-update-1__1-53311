VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00EDE7DC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  JPEG - Parser demo"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8970
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResult 
      BackColor       =   &H00E6FFE6&
      Height          =   3675
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   840
      Width           =   8640
   End
   Begin VB.CommandButton btnSelectFile 
      BackColor       =   &H00E0B4A3&
      Caption         =   " . . ."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   375
      Width           =   420
   End
   Begin VB.CommandButton btnGo 
      BackColor       =   &H006BD2FE&
      Caption         =   "Parse JPEG file"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7035
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1770
   End
   Begin VB.TextBox txtPathToJPEGfile 
      BackColor       =   &H00FAC5AD&
      Height          =   330
      Left            =   195
      TabIndex        =   0
      Text            =   ".\Smallest_JPEG.jpg"
      Top             =   375
      Width           =   5955
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "clsJPEGParser"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4125
      TabIndex        =   4
      Top             =   4725
      Width           =   1425
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "A demonstration of the class  clsJPEGParser   created in April 2004 by Light Templer. Thx for any comments (and votes ;) !)"
      Height          =   465
      Index           =   1
      Left            =   1650
      TabIndex        =   3
      Top             =   4725
      Width           =   5655
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the path to a JPEG file"
      Height          =   210
      Index           =   0
      Left            =   345
      TabIndex        =   1
      Top             =   135
      Width           =   2295
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00FDE4DB&
      BackStyle       =   1  'Opaque
      Height          =   600
      Left            =   195
      Shape           =   4  'Rounded Rectangle
      Top             =   4635
      Width           =   8610
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
'   frmMain.frm
'


' ======================================================================
' =  THIS FORM IS ONLY A SMALL DEMONSTARTION OF THE CLASS.             =
' =  ALL HINTS, COMMENTS AND CREDITS ARE IN FILE  clsJPEGparser.cls !  =
' ======================================================================



Option Explicit


' Used to get a path/filename by common dialog. Not neccessary for the class itself!
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" _
        (pOpenfilename As OPENFILENAME) As Long
        
Private Type OPENFILENAME
    lStructSize         As Long
    hWndOwner           As Long
    hInstance           As Long
    lpstrFilter         As String
    lpstrCustomFilter   As String
    nMaxCustFilter      As Long
    nFilterIndex        As Long
    lpstrFile           As String
    nMaxFile            As Long
    lpstrFileTitle      As String
    nMaxFileTitle       As Long
    lpstrInitialDir     As String
    lpstrTitle          As String
    flags               As Long
    nFileOffset         As Integer
    nFileExtension      As Integer
    lpstrDefExt         As String
    lCustData           As Long
    lpfnHook            As Long
    lpTemplateName      As String
End Type
'
'
'


Private Sub btnGo_Click()
    ' DEMO call to the class - Have a closer look to all the properties - thx!
    
    Dim oJPEGparser As clsJPEGparser
    
    txtResult.Text = ""
    
    Set oJPEGparser = New clsJPEGparser
    With oJPEGparser
    
        If .ParseJpegFile(txtPathToJPEGfile) = False Then
            txtResult.Text = vbCrLf + "Error parsing this file: " + vbCrLf + vbCrLf + " " + .ErrorMsg + _
            vbCrLf + vbCrLf + "====================================================" + vbCrLf + _
            " ---> Report until error occurs:" + vbCrLf + vbCrLf + .FullReport + vbCrLf
        Else
            txtResult.Text = .FullReport
        End If
        
    End With
    
    Set oJPEGparser = Nothing
    
End Sub





' ==================================================================
' From here to end it's only common stuff to handle
' 'Select a file' by API / common dialog. Just ignore it.
' ==================================================================

Private Sub btnSelectFile_Click()
    ' Select a JPEG file by common dialog
    
    Const sFilter As String = "JPEG files (*.Jpg)|*.Jpg|All Files (*.*)|*.*"

    Dim sFile As String

    sFile = OpenCommonDialog("Select a JPEG file to parse", sFilter) + vbNullChar
    sFile = Left$(sFile, InStr(1, sFile, Chr$(0)) - 1)
    If sFile <> "" Then
        txtPathToJPEGfile.Text = sFile
    End If

End Sub


Private Function OpenCommonDialog(Optional strDialogTitle As String = "Open", _
                                    Optional strFilter As String = "All Files|*.Jpg", _
                                    Optional strDefaultExtention As String = "*.*") As String
    
    Dim i               As Long
    Dim lLen            As Long
    Dim API_FileName    As OPENFILENAME
        
       
    With API_FileName
        .lpstrTitle = strDialogTitle
        .lpstrDefExt = strDefaultExtention
        
        .lpstrFilter = vbNullString
        lLen = Len(strFilter)
        For i = 1 To lLen
            If Mid(strFilter, i, 1) = "|" Then
                .lpstrFilter = .lpstrFilter + vbNullChar
            Else
                .lpstrFilter = .lpstrFilter + Mid(strFilter, i, 1)
            End If
        Next i
        .lpstrFilter = .lpstrFilter + vbNullChar
        
        .hInstance = App.hInstance
        .lpstrFile = vbNullChar & Space(259)
        .nMaxFile = 260
        .flags = &H4
        .lStructSize = Len(API_FileName)
        
        GetOpenFileName API_FileName        ' API call
        
        .lpstrFile = Trim(.lpstrFile)
        lLen = Len(.lpstrFile)
        If lLen <> 1 Then
            OpenCommonDialog = Trim(.lpstrFile)
        End If
    End With
    
End Function

' #*#

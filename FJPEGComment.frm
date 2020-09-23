VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "Comdlg32.ocx"
Begin VB.Form FJPEGComment 
   Caption         =   "Jpeg Comment"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWriteJpegInfo 
      Caption         =   "Write Jpeg Info"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmDial 
      Left            =   120
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "..."
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   480
      Width           =   4695
   End
   Begin VB.CommandButton cmdReadJpegInfo 
      Caption         =   "Read Jpeg Info"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtComment 
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   960
      Width           =   6015
   End
   Begin VB.CheckBox ChkVerbose 
      Caption         =   "Image info"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "File :"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "FJPEGComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strFileName As String
Private Sub cmdOpen_Click()
cmDial.Filter = "Jpeg Images (*.jpg;*.jpeg)|*.jpg;*.jpeg"
cmDial.ShowOpen
txtFileName.Text = cmDial.fileName

End Sub

Private Sub cmdReadJpegInfo_Click()

strFileName = txtFileName.Text
If strFileName = "" Then
    MsgBox "Select File first!", vbCritical
    Exit Sub
End If
If Dir(strFileName) = "" Then
    MsgBox "File not found!", vbCritical
    Exit Sub
End If



    txtComment.Text = scan_JPEG_header(strFileName, ChkVerbose.Value)

End Sub

Private Sub cmdWriteJpegInfo_Click()
strFileName = txtFileName.Text
If strFileName = "" Then
    MsgBox "Select File first!", vbCritical
    Exit Sub
End If
If Dir(strFileName) = "" Then
    MsgBox "File not found!", vbCritical
    Exit Sub
End If
Screen.MousePointer = vbHourglass
WriteJPGComment strFileName, txtComment.Text
'Like wrjpgcom.c - but it does not work, maybe I've omitted something
'write_JPEG_header strFileName, txtComment.Text
Screen.MousePointer = Normal
End Sub

Private Sub txtFileName_Change()
    txtComment.Text = ""
End Sub

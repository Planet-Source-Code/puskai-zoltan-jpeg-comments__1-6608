Attribute VB_Name = "MJPEGComment"
Public Const M_SOF0 = &HC0
Public Const M_SOF1 = &HC1
Public Const M_SOF2 = &HC2
Public Const M_SOF3 = &HC3
Public Const M_SOF5 = &HC5
Public Const M_SOF6 = &HC6
Public Const M_SOF7 = &HC7
Public Const M_SOF9 = &HC9
Public Const M_SOF10 = &HCA
Public Const M_SOF11 = &HCB
Public Const M_SOF13 = &HCD
Public Const M_SOF14 = &HCE
Public Const M_SOF15 = &HCF
Public Const M_SOI = &HD8
Public Const M_EOI = &HD9
Public Const M_SOS = &HDA
Public Const M_COM = &HFE
Public Const MAX_COM_LENGTH = 65000
Private fileNr As Integer
Private fileNrDest As Integer

Private Function NEXTBYTE() As Byte
Dim c As Byte
If EOF(fileNr) Then
    Exit Function
End If
Get #fileNr, , c
NEXTBYTE = c
End Function

Private Sub PUTBYTE(c As Byte)
Put #fileNrDest, , c
End Sub

Private Function read_1_byte() As Byte
Dim c As Byte
If EOF(fileNr) Then
    Exit Function
End If
Get #fileNr, , c
read_1_byte = c

End Function

Private Function read_2_byte() As Double
Dim c1 As Byte
Dim c2 As Byte
If EOF(fileNr) Then
    Exit Function
End If
Get #fileNr, , c1
If EOF(fileNr) Then
    Exit Function
End If
Get #fileNr, , c2
read_2_byte = CDbl(c1) * CDbl(256) + c2
End Function

Private Function next_marker() As Byte
Dim c As Byte
Dim discarded_bytes As Integer
c = read_1_byte
While c <> &HFF
    discarded_bytes = discarded_bytes + 1
    c = read_1_byte
Wend
Do
    c = read_1_byte
Loop While c = &HFF
'If discarded_bytes <> 0 Then MsgBox "Warning: garbage data found in JPEG file"
next_marker = c
End Function

Private Function first_marker() As Byte
Dim c1 As Byte, c2 As Byte
c1 = NEXTBYTE
c2 = NEXTBYTE
If c1 <> &HFF And c2 <> M_SOI Then MsgBox "not a JPEG file"
first_marker = c2
End Function
Private Sub skip_variable()
Dim length As Double
length = read_2_byte
If length < 2 Then MsgBox "Erroneous JPEG marker length"
length = length - 2
While length > 0
    read_1_byte
    length = length - 1
Wend
End Sub

Private Function process_COM() As String
Dim length  As Double
Dim ch As Byte
Dim s As String
lastch = 0
length = read_2_byte
If length < 2 Then MsgBox "Erroneous JPEG marker length"
length = length - 2
While length > 0
    ch = read_1_byte
    s = s & Chr(ch)
    length = length - 1
Wend
process_COM = s
End Function

Private Function process_SOFn(marker As Integer) As String
Dim length As Double
Dim image_height As Double, image_width As Double
Dim data_precision As Byte, num_components As Byte
Dim ci As Byte
Dim process As String

length = read_2_bytes
data_precision = read_1_byte
image_height = read_2_bytes
image_width = read_2_bytes
num_components = read_1_byte

Select Case marker
    Case M_SOF0:
        process = "Baseline"
    Case M_SOF1:
        process = "Extended sequential"
    Case M_SOF2:
        process = "Progressive"
    Case M_SOF3:
        process = "Lossless"
    Case M_SOF5:
        process = "Differential sequential"
    Case M_SOF6:
        process = "Differential progressive"
    Case M_SOF7:
        process = "Differential lossless"
    Case M_SOF9:
        process = "Extended sequential, arithmetic coding"
    Case M_SOF10:
        process = "Progressive, arithmetic coding"
    Case M_SOF11:
        process = "Lossless, arithmetic coding"
    Case M_SOF13:
        process = "Differential sequential, arithmetic coding"
    Case M_SOF14:
        process = "Differential progressive, arithmetic coding"
    Case M_SOF15:
        process = "Differential lossless, arithmetic coding"
    Case Else:
        process = "Unknown"
End Select
process_SOFn = "JPEG image is " & image_width & " * " & image_height & " ," & num_components & " color components ," & data_precision & " bits per sample"
If length <> (8 + CDbl(num_components) * 3) Then MsgBox "Bogus SOF marker length"
For ci = 0 To num_components - 1
    read_1_byte  '   Component ID code
    read_1_byte ' H, V sampling factors
    read_1_byte ' Quantization table number
Next
End Function



Private Sub write_1_byte(c As Byte)
     PUTBYTE c
End Sub

Private Sub write_2_byte(c As Double)
    PUTBYTE ((c / 256) And &HFF)
    PUTBYTE (c And &HFF)
End Sub

Private Sub write_marker(marker As Byte)
    PUTBYTE &HFF
    PUTBYTE marker
End Sub

Private Sub copy_rest_of_file()
While Not EOF(fileNr)
    PUTBYTE NEXTBYTE
Wend
End Sub

Private Sub copy_variable()
Dim length As Double
length = read_2_byte
write_2_byte length
'If (length < 2) Then MsgBox "Erroneous JPEG marker length", vbCritical
length = length - 2
While (length > 0)
    write_1_byte read_1_byte
length = length - 1
Wend

End Sub

Public Static Function write_JPEG_header(strFileName As String, strComment As String) As String
Dim maker As Integer
Dim strdestFile As String
fileNr = FreeFile
Open strFileName For Binary Access Read As #fileNr
strdestFile = strFileName & "temp.jpg"
fileNrDest = FreeFile
Open strFileName For Random As #fileNrDest

'If first_marker <> M_SOI Then MsgBox "Expected SOI marker first", vbCritical
write_marker M_SOI
Do
    marker = next_marker
    Select Case marker
        Case M_SOF0:        ' Baseline
        Case M_SOF1:        ' Extended sequential, Huffman
        Case M_SOF2:        ' Progressive, Huffman
        Case M_SOF3:        ' Lossless, Huffman
        Case M_SOF5:        ' Differential sequential, Huffman
        Case M_SOF6:        ' Differential progressive, Huffman
        Case M_SOF7:        ' Differential lossless, Huffman
        Case M_SOF9:        ' Extended sequential, arithmetic
        Case M_SOF10:       ' Progressive, arithmetic
        Case M_SOF11:       ' Lossless, arithmetic
        Case M_SOF13:       ' Differential sequential, arithmetic
        Case M_SOF14:       ' Differential progressive, arithmetic
        Case M_SOF15:       ' Differential lossless, arithmetic
        Case M_SOS:         ' should not see compressed data before SOF
            'MsgBox "SOS without prior SOF"
        Case M_EOI:         ' in case it's a tables-only JPEG stream
        Case M_COM:         ' Existing COM: conditionally discard
            If strComment <> "" Then
                write_marker (marker)
                copy_variable
             Else
                skip_variable
            End If
        Case Else
            write_marker (marker)
            'copy_variable      ' we assume it has a parameter count...
    End Select
Loop
Close #fileNr
Close #fileNrDest
End Function

Public Function scan_JPEG_header(strFileName As String, verbose As Boolean) As String
Dim marker As Integer
Dim returnedString As String

fileNr = FreeFile
Open strFileName For Binary Access Read As #fileNr

If first_marker <> M_SOI Then
    MsgBox "Expected SOI marker first"
    Exit Function
End If
Do
    marker = next_marker
    Select Case marker
        Case M_SOF0:        ' Baseline
        Case M_SOF1:        ' Extended sequential, Huffman
        Case M_SOF2:        ' Progressive, Huffman
        Case M_SOF3:        ' Lossless, Huffman
        Case M_SOF5:        ' Differential sequential, Huffman
        Case M_SOF6:        ' Differential progressive, Huffman
        Case M_SOF7:        ' Differential lossless, Huffman
        Case M_SOF9:        ' Extended sequential, arithmetic
        Case M_SOF10:       ' Progressive, arithmetic
        Case M_SOF11:       ' Lossless, arithmetic
        Case M_SOF13:       ' Differential sequential, arithmetic
        Case M_SOF14:       ' Differential progressive, arithmetic
        Case M_SOF15:       ' Differential lossless, arithmetic
            If verbose Then
                returnedString = returnedString & process_SOFn(marker) & vbTab
            Else
                skip_variable
            End If
        Case M_SOS:         ' stop before hitting compressed data
            scan_JPEG_header = returnedString
            Close #fileNr
            Exit Function
        Case M_EOI:         ' in case it's a tables-only JPEG stream
            scan_JPEG_header = returnedString
            Close #fileNr
            Exit Function
        Case M_COM:
            returnedString = returnedString & process_COM & vbTab
        Case Else:            ' Anything else just gets skipped
            skip_variable ' we assume it has a parameter count...
        End Select
Loop
Close #fileNr
End Function



Public Sub WriteJPGComment(fileName As String, comment As String)
Dim notyet As Boolean
Dim fileNametemp As String
Dim a As Byte, s1 As Byte, s2 As Byte, l1 As Byte, l2 As Byte, t As Byte, x As Byte
Dim i As Long
Dim sComment As String
Dim fileNr As Integer
Dim fileNrTemp As Integer
Dim FilePos As Double
Dim FilePosTemp As Double
Dim MySize As Double
Dim N As Integer


fileNametemp = fileName & "temp.jpg"
'On Error Resume Next
If Dir(fileNametemp) <> "" Then
    Kill fileNametemp
End If
notyet = True
fileNr = FreeFile
Open fileName For Binary As #fileNr
fileNrTemp = FreeFile
Open fileNametemp For Binary Access Write As #fileNrTemp


Get #fileNr, , a
Put #fileNrTemp, , a
Get #fileNr, , a
Put #fileNrTemp, , a
Do
    Get #fileNr, , t
    Put #fileNrTemp, , t
    While t <> &HFF
        Get #fileNr, , t
        Put #fileNrTemp, , t
    Wend
    s1 = &HFF
    Get #fileNr, , s2
        While s2 = &HFF
            Get #fileNr, , s2
        Wend
    Get #fileNr, , l1
    Get #fileNr, , l2
    sComment = ""
    For i = 1 To (CDbl(256) * CDbl(l1) + l2) - 2
        If Not EOF(fileNr) Then
                Get #fileNr, , x
                sComment = sComment & Chr(x)
        End If
    Next
    If ((s2 And &HF0) = &HC0) And notyet And (comment <> "") Then
        a = &HFF
        Put #fileNrTemp, , a
        a = &HFE
        Put #fileNrTemp, , a
        a = CByte((Len(comment) + 2) / 256)
        Put #fileNrTemp, , a
        a = CByte((Len(comment) + 2) Mod 256)
        Put #fileNrTemp, , a
        For i = 1 To Len(comment)
            a = CByte(Asc(Mid(comment, i, 1)))
            Put #fileNrTemp, , a
        Next
        notyet = False
    End If
    If s2 <> &HFE Then
        Put #fileNrTemp, , s1
        Put #fileNrTemp, , s2
        Put #fileNrTemp, , l1
        Put #fileNrTemp, , l2
        For i = 1 To Len(sComment)
            a = CByte(Asc(Mid(sComment, i, 1)))
            Put #fileNrTemp, , a
        Next
        
    End If
Loop Until EOF(fileNr) Or s2 = &HDA


'to awoid
'        While Not EOF(fileNr)
'            Get #fileNr, , a
'            Put #fileNrTemp, , a
'        Wend
'       Close #fileNr
'       Close #fileNrTemp
'       Kill fileName
'       Name fileNametemp As fileName
'       End Sub

FilePos = Seek(fileNr)
FilePosTemp = Seek(fileNrTemp)
MySize = LOF(fileNr) - FilePos
ReDim arr(MySize)
Close #fileNr
Close #fileNrTemp
fileNr = FreeFile
Open fileName For Binary As #fileNr
fileNrTemp = FreeFile
Open fileNametemp For Binary As #fileNrTemp
Seek #fileNr, FilePos
Seek #fileNrTemp, FilePosTemp
N = Int(MySize / 4096)
If N Then
    For i = 1 To N
        b$ = Space$(4096)
        Get #fileNr, , b$
        Put #fileNrTemp, , b$
    Next
End If
N = MySize Mod 4096
If N Then
    b$ = Space$(N)
    Get #fileNr, , b$
    Put #fileNrTemp, , b$
End If
Close #fileNr
Close #fileNrTemp
'I have commented this not to loose some picture
'Kill fileName
'Name fileNametemp As fileName
End Sub



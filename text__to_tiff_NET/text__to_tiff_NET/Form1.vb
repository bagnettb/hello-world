Imports System.IO

Public Class Form1





    Private Sub Command1_Click(sender As Object, e As EventArgs) Handles Command1.Click

    End Sub


    Const DWordSz = 32
    Const WordSz = 16
    Const DataOrg = 62

    Dim ifdblock As Object  ' IFD BLOCK
    Dim ImgSz(4) As Byte
    Dim ImgWd(4) As Byte
    Dim ImgHt(4) As Byte
    Dim LWord(4) As Byte
    Dim BDaSz(4) As Byte
    Dim bmp(4000000) As Byte
    Dim bsz As Long
    Dim bmapsz As Long
    Dim bmaprowsz As Long
    Dim InLine As String

    Dim w As Long
    Dim h As Long

    Dim txtLine As Long
    Dim txtCol As Long

    Dim arow As Object
    Dim acol As Object
    Dim brow As Object
    Dim bcol As Object
    Dim crow As Object
    Dim ccol As Object
    Dim erow As Object
    Dim ecol As Object
    Dim frow As Object
    Dim fcol As Object
    Dim grow As Object
    Dim gcol As Object

    Dim hrow As Object
    Dim hcol As Object
    Dim irow As Object
    Dim icol As Object
    Dim rrow As Object
    Dim rcol As Object
    Dim nrow As Object
    Dim ncol As Object
    Dim orow As Object
    Dim ocol As Object
    Dim prow As Object
    Dim pcol As Object
    Dim qrow As Object
    Dim qcol As Object
    Dim trow As Object
    Dim tcol As Object
    Dim urow As Object
    Dim ucol As Object
    Dim yrow As Object
    Dim ycol As Object
    Dim drow As Object
    Dim dcol As Object
    Dim jrow As Object
    Dim jcol As Object
    Dim krow As Object
    Dim kcol As Object
    Dim lrow As Object
    Dim lcol As Object
    Dim mrow As Object
    Dim mcol As Object
    Dim srow As Object
    Dim scol As Object
    Dim vrow As Object
    Dim vcol As Object
    Dim wrow As Object
    Dim wcol As Object
    Dim xrow As Object
    Dim xcol As Object
    Dim zrow As Object
    Dim zcol As Object
    Dim row1 As Object
    Dim col1 As Object
    Dim row2 As Object
    Dim col2 As Object
    Dim row3 As Object
    Dim col3 As Object
    Dim row4 As Object
    Dim col4 As Object
    Dim row5 As Object
    Dim col5 As Object
    Dim row6 As Object
    Dim col6 As Object
    Dim row7 As Object
    Dim col7 As Object
    Dim row8 As Object
    Dim col8 As Object
    Dim row9 As Object
    Dim col9 As Object
    Dim row0 As Object
    Dim col0 As Object
    Dim s1row As Object
    Dim s1col As Object
    Dim s2row As Object
    Dim s2col As Object
    Dim s3row As Object
    Dim s3col As Object
    Dim s4row As Object
    Dim s4col As Object
    Dim s5row As Object
    Dim s5col As Object
    Dim s6row As Object
    Dim s6col As Object
    Dim s7row As Object
    Dim s7col As Object
    Dim s8row As Object
    Dim s8col As Object
    Dim s9row As Object
    Dim s9col As Object
    Dim s10row As Object
    Dim s10col As Object
    Dim s11row As Object
    Dim s11col As Object
    Dim s12row As Object
    Dim s12col As Object
    Dim s13row As Object
    Dim s13col As Object
    Dim s14row As Object
    Dim s14col As Object
    Dim s15row As Object
    Dim s15col As Object
    Dim s16row As Object
    Dim s16col As Object
    Dim s17row As Object
    Dim s17col As Object
    Dim s18row As Object
    Dim s18col As Object
    Dim s19row As Object
    Dim s19col As Object
    Dim s20row As Object
    Dim s20col As Object
    Dim s21row As Object
    Dim s21col As Object
    Dim s22row As Object
    Dim s22col As Object
    Dim s23row As Object
    Dim s23col As Object
    Dim s24row As Object
    Dim s24col As Object
    Dim s25row As Object
    Dim s25col As Object
    Dim s26row As Object
    Dim s26col As Object
    Dim s27row As Object
    Dim s27col As Object
    Dim s28row As Object
    Dim s28col As Object
    Dim s29row As Object
    Dim s29col As Object
    Dim s30row As Object
    Dim s30col As Object
    Dim s31row As Object
    Dim s31col As Object
    Dim s32row As Object
    Dim s32col As Object



    '
    ' ----- < Command1_Click > -----
    '
    Private Sub Command1_Click()
        Call LoadLetters()
        Call bits2DWords(34)

        Call startBMP()
        End

    End Sub



    '
    ' ----- < resetbmp > -----
    '
    Sub resetbmp()
        bsz = UBound(bmp)
        For i = 1 To bsz
            bmp(i) = &H0
        Next i
    End Sub

    '
    ' ----- < long2DWord > -----
    ' This subroutine load the hex value in backwards as opposed to long2word

    ' which element in the array do you wish to copy the bytes?
    Private Sub long2DWord(p_in As Long, p_ix As Long)
        Dim k As Long
        Dim qu As Long
        Dim ix As Long
        Dim b As Byte
        ' subtract .4999999 so that it truncates instead of rounds


        qu = p_in
        ix = p_ix

        For k = 0 To 3 Step 1
            bmp(ix + k) = qu Mod 256
            If qu >= 256 Then
                qu = CLng((qu / 256) - 0.4999999)
            Else
                qu = 0
            End If
        Next k

    End Sub

    '
    ' ----- < long2BackWord > -----
    '
    ' which element in the array do you wish to copy the bytes?
    Private Sub long2BackWord(p_in As Long, p_ix)
        Dim k As Long
        Dim qu As Long
        Dim b As Byte
        Dim ix As Long
        ' subtract .4999999 so that it truncates instead of rounds

        qu = p_in
        ix = p_ix

        For k = 0 To 1 Step 1
            bmp(ix + k) = qu Mod 256
            If qu >= 256 Then
                qu = CLng((qu / 256) - 0.4999999)
            Else
                qu = 0
            End If
        Next k

    End Sub


    '
    ' ----- < long2Word > -----
    '
    ' which element in the array do you wish to copy the bytes?
    Private Sub long2Word(p_in As Long, p_ix)
        Dim k As Long
        Dim qu As Long
        Dim ix As Long
        Dim b As Byte
        ' subtract .4999999 so that it truncates instead of rounds

        qu = p_in
        ix = p_ix

        For k = 1 To 0 Step -1
            bmp(ix + k) = qu Mod 256
            If qu >= 256 Then
                qu = CLng((qu / 256) - 0.4999999)
            Else
                qu = 0
            End If
        Next k

    End Sub



    '
    ' ----- < byteop > -----
    '
    Private Function byteop(p_bit As Long, p_posn As Long, p_byte As Byte) As Byte
        ' 00...FF
        ' take input byte, perform operation on it and export result byte
        Dim bite As Byte

        If p_bit = 1 Then
            Select Case p_posn
                Case 0
                    bite = p_byte Or &H80
                Case 1
                    bite = p_byte Or &H40
                Case 2
                    bite = p_byte Or &H20
                Case 3
                    bite = p_byte Or &H10
                Case 4
                    bite = p_byte Or &H8
                Case 5
                    bite = p_byte Or &H4
                Case 6
                    bite = p_byte Or &H2
                Case 7
                    bite = p_byte Or &H1
            End Select
        Else
            Select Case p_posn
                Case 0
                    bite = p_byte And &H7F
                Case 1
                    bite = p_byte And &HBF
                Case 2
                    bite = p_byte And &HDF
                Case 3
                    bite = p_byte And &HEF
                Case 4
                    bite = p_byte And &HF7
                Case 5
                    bite = p_byte And &HFB
                Case 6
                    bite = p_byte And &HFD
                Case 7
                    bite = p_byte And &HFE
            End Select
        End If

        byteop = bite
    End Function

    '
    ' ----- < bits2DWords > -----
    ' buffer bits into DWords and xfer them to array bmp()
    Private Sub bits2DWords(p_bitcnt As Long)
        Dim bitsleft As Long

        bitsleft = p_bitcnt
        ' i = RoundDw(p_bitcnt)
        While bitsleft >= 32
            ' xfer full DWord to array
            Call DWord2BMP(bsz, &HFF, &HFF, &HFF, &HFF)
            bitsleft = bitsleft - 32
Wend
' handle partial DWord.  Numbers are zero-relative
Select Case bitsleft
            Case 0 : Call DWord2BMP(bsz, &H80, &H0, &H0, &H0)
            Case 1 : Call DWord2BMP(bsz, &HA0, &H0, &H0, &H0)
            Case 2 : Call DWord2BMP(bsz, &HC0, &H0, &H0, &H0)
            Case 3 : Call DWord2BMP(bsz, &HF0, &H0, &H0, &H0)
            Case 4 : Call DWord2BMP(bsz, &HF8, &H0, &H0, &H0)
            Case 5 : Call DWord2BMP(bsz, &HFA, &H0, &H0, &H0)
            Case 6 : Call DWord2BMP(bsz, &HFC, &H0, &H0, &H0)
            Case 7 : Call DWord2BMP(bsz, &HFF, &H0, &H0, &H0)
            Case 8 : Call DWord2BMP(bsz, &HFF, &H80, &H0, &H0)
            Case 9 : Call DWord2BMP(bsz, &HFF, &HA0, &H0, &H0)
            Case 10 : Call DWord2BMP(bsz, &HFF, &HC0, &H0, &H0)
            Case 11 : Call DWord2BMP(bsz, &HFF, &HF0, &H0, &H0)
            Case 12 : Call DWord2BMP(bsz, &HFF, &HF8, &H0, &H0)
            Case 13 : Call DWord2BMP(bsz, &HFF, &HFA, &H0, &H0)
            Case 14 : Call DWord2BMP(bsz, &HFF, &HFC, &H0, &H0)
            Case 15 : Call DWord2BMP(bsz, &HFF, &HFF, &H0, &H0)
            Case 16 : Call DWord2BMP(bsz, &HFF, &HFF, &H80, &H0)
            Case 17 : Call DWord2BMP(bsz, &HFF, &HFF, &HA0, &H0)
            Case 18 : Call DWord2BMP(bsz, &HFF, &HFF, &HC0, &H0)
            Case 19 : Call DWord2BMP(bsz, &HFF, &HFF, &HF0, &H0)
            Case 20 : Call DWord2BMP(bsz, &HFF, &HFF, &HF8, &H0)
            Case 21 : Call DWord2BMP(bsz, &HFF, &HFF, &HFA, &H0)
            Case 22 : Call DWord2BMP(bsz, &HFF, &HFF, &HFC, &H0)
            Case 23 : Call DWord2BMP(bsz, &HFF, &HFF, &HFF, &H0)
            Case 24 : Call DWord2BMP(bsz, &HFF, &HFF, &HFF, &H80)
            Case 25 : Call DWord2BMP(bsz, &HFF, &HFF, &HFF, &HA0)
            Case 26 : Call DWord2BMP(bsz, &HFF, &HFF, &HFF, &HC0)
            Case 27 : Call DWord2BMP(bsz, &HFF, &HFF, &HFF, &HF0)
            Case 25 : Call DWord2BMP(bsz, &HFF, &HFF, &HFF, &HF8)
            Case 29 : Call DWord2BMP(bsz, &HFF, &HFF, &HFF, &HFA)
            Case 30 : Call DWord2BMP(bsz, &HFF, &HFF, &HFF, &HFC)
            Case 31 : Call DWord2BMP(bsz, &HFF, &HFF, &HFF, &HFF)
        End Select
    End Sub

    '
    ' ----- < bits2Decimal > -----
    '
    Private Function bits2Decimal(p_in As String) As Long
        ' change bit patterns stored in string to decimal number
        Dim result As Long
        Dim bitstr As String
        Dim currchar As Char

        ' bitstr = "1111111100000000"
        bitstr = p_in

        result = 0
        For i = 1 To WordSz
            currchar = Mid(bitstr, i, 1)
            If currchar = "1" Then
                result = (result * 2) + 1
            Else
                result = result * 2
            End If
        Next i
        bits2Decimal = result

    End Function

    '
    ' ----- < DWord2BMP > -----
    '
    Private Sub DWord2BMP(p_in As Long, b3 As Byte, b2 As Byte, b1 As Byte, b0 As Byte)
        ix = p_in
        bsz = bsz + 1 : bmp(bsz) = b3
        bsz = bsz + 1 : bmp(bsz) = b2
        bsz = bsz + 1 : bmp(bsz) = b1
        bsz = bsz + 1 : bmp(bsz) = b0

    End Sub

    '
    ' ----- < writeBMP > -----
    '  https://www.daniweb.com/programming/software-development/threads/311504/how-to-convert-get-put-in-vb-net
    '

    Private Sub writeBMP()


        Dim FileStream = New FileStream("tifflog.bmp", FileMode.Create)
        FileStream.Seek(0, SeekOrigin.Begin)

        FileStream.Close()




        Dim ix As Long
        ix = 0
        On Error GoTo handle
        While ix <= bsz And ix <= UBound(bmp)
            'Put #34, , CByte(bmp(ix))
            FileStream.WriteByte(CByte(bmp(ix)))
            ix = ix + 1
        End While
        Exit Sub
handle:
        MsgBox("There was an Error writing To the file.   Is it open In another application?")
    End Sub

    '
    ' ----- < WriteBitmap > -----
    '
    Private Sub WriteBitmap(p_width As Long, p_height As Long)
        Dim cols As Long
        ' round # of columns up to nearest DWord (32 bits)
        cols = (p_width + 16) / 32
    End Sub

    '
    ' ----- < RoundDw > -----
    ' Determines how many DWords you need to
    '  hold the number of pixel columns
    Private Function RoundDw(p_in As Long) As Long
        If (p_in / DWordSz - CLng(p_in / DWordSz)) > 0 Then
            RoundDw = CLng(p_in / DWordSz) + 1
        Else
            RoundDw = CLng(p_in / DWordSz)
        End If
    End Function  '  RoundDw

    Private Sub drawbit(p_row As Long, p_col As Long)
        ' add .4999999 so that it rounds up
        Dim colbit As Long
        Dim colbyte As Long
        Dim posn As Long

        colbyte = CLng(((p_col + 1) / 8) + 0.4999999) - 1
        colbit = p_col Mod 8
        posn = (DataOrg + bmaprowsz * p_row) + colbyte
        ' byteop(p_bit As Long, p_posn As Long, p_byte As Byte)
        bmp(posn) = byteop(0, colbit, bmp(posn))

    End Sub


    Private Sub DrawLetter(p_row As Object, p_col As Object, p_rorg As Long, p_corg As Long)
        Dim absrow As Long
        Dim abscol As Long
        Dim EOArr As Long

        EOArr = UBound(p_row)
        For i = 0 To EOArr
            absrow = CLng(p_row(i) + p_rorg)
            abscol = CLng(p_col(i) + p_corg)
            Call drawbit(absrow, abscol)

        Next i
    End Sub

    Private Sub DrawWord(p_line As Long, p_col As Long, p_str As String)
        Dim spacer As Long
        Dim strLen As Long
        Dim ltr As String

        spacer = 11
        strLen = Len(p_str)

        For i = 1 To strLen
            ltr = UCase(Mid(p_str, i, 1))
            Select Case ltr
                Case " "
                    txtCol = txtCol + 12
                Case "A"
                    Call DrawLetter(arow, acol, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "B"
                    Call DrawLetter(brow, bcol, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "C"
                    Call DrawLetter(crow, ccol, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "D"
                    Call DrawLetter(drow, dcol, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "E"
                    Call DrawLetter(erow, ecol, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "F"
                    Call DrawLetter(frow, fcol, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "G"
                    Call DrawLetter(grow, gcol, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "H"
                    Call DrawLetter(hrow, hcol, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "I"
                    Call DrawLetter(irow, icol, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "J"
                    Call DrawLetter(jrow, jcol, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "K"
                    Call DrawLetter(krow, kcol, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "L"
                    Call DrawLetter(lrow, lcol, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "M"
                    Call DrawLetter(mrow, mcol, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "N"
                    Call DrawLetter(nrow, ncol, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "O"
                    Call DrawLetter(orow, ocol, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "P"
                    Call DrawLetter(prow, pcol, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "Q"
                    Call DrawLetter(qrow, qcol, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "R"
                    Call DrawLetter(rrow, rcol, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "S"
                    Call DrawLetter(srow, scol, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "T"
                    Call DrawLetter(trow, tcol, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "U"
                    Call DrawLetter(urow, ucol, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "V"
                    Call DrawLetter(vrow, vcol, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "W"
                    Call DrawLetter(wrow, wcol, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "X"
                    Call DrawLetter(xrow, xcol, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "Y"
                    Call DrawLetter(yrow, ycol, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "Z"
                    Call DrawLetter(zrow, zcol, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "1"
                    Call DrawLetter(row1, col1, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "2"
                    Call DrawLetter(row2, col2, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "3"
                    Call DrawLetter(row3, col3, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "4"
                    Call DrawLetter(row4, col4, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "5"
                    Call DrawLetter(row5, col5, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "6"
                    Call DrawLetter(row6, col6, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "7"
                    Call DrawLetter(row7, col7, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "8"
                    Call DrawLetter(row8, col8, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "9"
                    Call DrawLetter(row9, col9, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "0"
                    Call DrawLetter(row0, col0, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "-"
                    Call DrawLetter(s1row, s1col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "="
                    Call DrawLetter(s2row, s2col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "+"
                    Call DrawLetter(s3row, s3col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "_"
                    Call DrawLetter(s4row, s4col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "*"
                    Call DrawLetter(s5row, s5col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "("
                    Call DrawLetter(s6row, s6col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case ")"
                    Call DrawLetter(s7row, s7col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "\"
                    Call DrawLetter(s8row, s8col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "/"
                    Call DrawLetter(s9row, s9col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "'"
                    Call DrawLetter(s10row, s10col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case """"
                    Call DrawLetter(s11row, s11col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "|"
                    Call DrawLetter(s12row, s12col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "<"
                    Call DrawLetter(s13row, s13col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case ">"
                    Call DrawLetter(s14row, s14col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "["
                    Call DrawLetter(s15row, s15col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "]"
                    Call DrawLetter(s16row, s16col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case ";"
                    Call DrawLetter(s17row, s17col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case ":"
                    Call DrawLetter(s18row, s18col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "."
                    Call DrawLetter(s19row, s19col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "?"
                    Call DrawLetter(s20row, s20col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "&"
                    Call DrawLetter(s21row, s21col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case ","
                    Call DrawLetter(s22row, s22col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "!"
                    Call DrawLetter(s23row, s23col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "^"
                    Call DrawLetter(s24row, s24col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "@"
                    Call DrawLetter(s25row, s25col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "#"
                    Call DrawLetter(s26row, s26col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "$"
                    Call DrawLetter(s27row, s27col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "%"
                    Call DrawLetter(s28row, s28col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "{"
                    Call DrawLetter(s29row, s29col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "}"
                    Call DrawLetter(s30row, s30col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "~"
                    Call DrawLetter(s31row, s31col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case "`"
                    Call DrawLetter(s32row, s32col, txtLine, txtCol) : txtCol = txtCol + spacer
                Case Else
                    ' draw underscore
                    Call DrawLetter(s4row, s4col, txtLine, txtCol) : txtCol = txtCol + spacer
            End Select
        Next i
        ' -=+_*()\/  '"|<>[];:.?&
    End Sub

    '
    ' ----- < LoadLetters > -----
    '
    Private Sub LoadLetters()
        Dim arow = New Integer() {1, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 6, 6, 6, 6, 7, 7, 8, 8, 9, 9, 9, 9, 9, 9}
        Dim acol = New Integer() {4, 5, 6, 5, 6, 4, 7, 4, 7, 4, 7, 3, 4, 5, 6, 7, 8, 3, 8, 2, 9, 1, 2, 3, 8, 9, 10}
        Dim brow = New Integer() {1, 1, 1, 1, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 5, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 9, 9, 9, 9}
        Dim bcol = New Integer() {1, 2, 3, 4, 5, 6, 2, 7, 2, 7, 2, 7, 2, 3, 4, 5, 6, 2, 7, 2, 7, 2, 7, 1, 2, 3, 4, 5, 6}
        Dim crow = New Integer() {1, 1, 1, 1, 1, 2, 2, 2, 3, 3, 4, 5, 6, 7, 8, 8, 9, 9, 9, 9, 9}
        Dim ccol = New Integer() {3, 4, 5, 6, 8, 2, 7, 8, 1, 8, 1, 1, 1, 1, 2, 8, 3, 4, 5, 6, 7}
        Dim erow = New Integer() {1, 1, 1, 1, 1, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 9, 9, 9, 9, 9}
        Dim ecol = New Integer() {1, 2, 3, 4, 5, 6, 7, 2, 7, 2, 7, 2, 5, 2, 3, 4, 5, 2, 5, 2, 7, 2, 7, 1, 2, 3, 4, 5, 6, 7}
        Dim frow = New Integer() {1, 1, 1, 1, 1, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 5, 5, 6, 6, 7, 8, 9, 9, 9, 9, 9}
        Dim fcol = New Integer() {1, 2, 3, 4, 5, 6, 7, 2, 7, 2, 7, 2, 5, 2, 3, 4, 5, 2, 5, 2, 2, 1, 2, 3, 4, 5}
        Dim grow = New Integer() {1, 1, 1, 1, 1, 2, 2, 2, 3, 4, 5, 6, 6, 6, 6, 6, 6, 7, 7, 8, 8, 9, 9, 9, 9, 9}
        Dim gcol = New Integer() {3, 4, 5, 6, 8, 2, 7, 8, 1, 1, 1, 1, 5, 6, 7, 8, 9, 1, 8, 2, 8, 3, 4, 5, 6, 7}
        Dim hrow = New Integer() {1, 1, 1, 1, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 5, 5, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 9, 9, 9, 9}
        Dim hcol = New Integer() {1, 2, 3, 6, 7, 8, 2, 7, 2, 7, 2, 7, 2, 3, 4, 5, 6, 7, 2, 7, 2, 7, 2, 7, 1, 2, 3, 6, 7, 8}
        Dim irow = New Integer() {1, 1, 1, 1, 1, 1, 1, 2, 3, 4, 5, 6, 7, 8, 9, 9, 9, 9, 9, 9, 9}
        Dim icol = New Integer() {1, 2, 3, 4, 5, 6, 7, 4, 4, 4, 4, 4, 4, 4, 1, 2, 3, 4, 5, 6, 7}
        Dim rrow = New Integer() {1, 1, 1, 1, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 5, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 9, 9, 9}
        Dim rcol = New Integer() {1, 2, 3, 4, 5, 6, 2, 7, 2, 7, 2, 7, 2, 3, 4, 5, 6, 2, 5, 2, 6, 2, 7, 1, 2, 3, 7, 8}
        Dim nrow = New Integer() {1, 1, 1, 1, 1, 1, 2, 2, 2, 3, 3, 3, 4, 4, 4, 5, 5, 5, 6, 6, 6, 7, 7, 7, 8, 8, 8, 9, 9, 9, 9, 9}
        Dim ncol = New Integer() {1, 2, 3, 7, 8, 9, 2, 3, 8, 2, 4, 8, 2, 4, 8, 2, 5, 8, 2, 6, 8, 2, 6, 8, 2, 7, 8, 1, 2, 3, 7, 8}
        Dim orow = New Integer() {1, 1, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 9, 9}
        Dim ocol = New Integer() {3, 4, 5, 6, 2, 7, 1, 8, 1, 8, 1, 8, 1, 8, 1, 8, 2, 7, 3, 4, 5, 6}
        Dim prow = New Integer() {1, 1, 1, 1, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 6, 6, 6, 7, 8, 9, 9, 9, 9, 9}
        Dim pcol = New Integer() {1, 2, 3, 4, 5, 6, 2, 7, 2, 7, 2, 7, 2, 7, 2, 3, 4, 5, 6, 2, 2, 1, 2, 3, 4, 5}
        Dim qrow = New Integer() {1, 1, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 9, 9, 10, 10, 10, 10, 10, 11}
        Dim qcol = New Integer() {3, 4, 5, 6, 2, 7, 1, 8, 1, 8, 1, 8, 1, 8, 1, 8, 2, 7, 3, 4, 5, 6, 4, 5, 6, 7, 8, 3}
        Dim trow = New Integer() {1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 3, 3, 3, 4, 4, 4, 5, 6, 7, 8, 9, 9, 9, 9, 9}
        Dim tcol = New Integer() {1, 2, 3, 4, 5, 6, 7, 1, 4, 7, 1, 4, 7, 1, 4, 7, 4, 4, 4, 4, 2, 3, 4, 5, 6}
        Dim urow = New Integer() {1, 1, 1, 1, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 9, 9}
        Dim ucol = New Integer() {1, 2, 3, 6, 7, 8, 2, 7, 2, 7, 2, 7, 2, 7, 2, 7, 2, 7, 2, 7, 3, 4, 5, 6}
        Dim yrow = New Integer() {1, 1, 1, 1, 1, 1, 2, 2, 3, 3, 4, 4, 5, 6, 7, 8, 9, 9, 9, 9, 9}
        Dim ycol = New Integer() {1, 2, 3, 7, 8, 9, 2, 8, 3, 7, 4, 6, 5, 5, 5, 5, 3, 4, 5, 6, 7}
        Dim drow = New Integer() {1, 1, 1, 1, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 9, 9, 9, 9}
        Dim dcol = New Integer() {1, 2, 3, 4, 5, 6, 2, 7, 2, 8, 2, 8, 2, 8, 2, 8, 2, 8, 2, 7, 1, 2, 3, 4, 5, 6}
        Dim jrow = New Integer() {1, 1, 1, 1, 1, 1, 2, 3, 4, 5, 6, 6, 7, 7, 8, 8, 9, 9, 9, 9}
        Dim jcol = New Integer() {3, 4, 5, 6, 7, 8, 6, 6, 6, 6, 1, 6, 1, 6, 1, 6, 2, 3, 4, 5}
        Dim krow = New Integer() {1, 1, 1, 1, 1, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 9, 9, 9}
        Dim kcol = New Integer() {1, 2, 3, 5, 6, 7, 8, 2, 6, 2, 5, 2, 4, 2, 3, 4, 5, 2, 6, 2, 6, 2, 7, 1, 2, 3, 7, 8}
        Dim lrow = New Integer() {1, 1, 1, 1, 1, 2, 3, 4, 5, 6, 6, 7, 7, 8, 8, 9, 9, 9, 9, 9, 9, 9, 9}
        Dim lcol = New Integer() {1, 2, 3, 4, 5, 3, 3, 3, 3, 3, 8, 3, 8, 3, 8, 1, 2, 3, 4, 5, 6, 7, 8}
        Dim mrow = New Integer() {1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 6, 6, 6, 7, 7, 8, 8, 9, 9, 9, 9, 9, 9}
        Dim mcol = New Integer() {1, 2, 3, 7, 8, 9, 2, 3, 7, 8, 2, 4, 6, 8, 2, 4, 6, 8, 2, 4, 6, 8, 2, 5, 8, 2, 8, 2, 8, 1, 2, 3, 7, 8, 9}
        Dim srow = New Integer() {1, 1, 1, 1, 2, 2, 2, 3, 3, 4, 5, 5, 5, 5, 6, 7, 7, 8, 8, 8, 9, 9, 9, 9}
        Dim scol = New Integer() {2, 3, 4, 6, 1, 5, 6, 1, 6, 1, 2, 3, 4, 5, 6, 1, 6, 1, 2, 6, 1, 3, 4, 5}
        Dim vrow = New Integer() {1, 1, 1, 1, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9}
        Dim vcol = New Integer() {1, 2, 3, 8, 9, 10, 2, 9, 3, 8, 3, 8, 4, 7, 4, 7, 4, 7, 5, 6, 5, 6}
        Dim wrow = New Integer() {1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 3, 3, 3, 4, 4, 4, 5, 5, 5, 5, 6, 6, 6, 6, 7, 7, 7, 7, 8, 8, 8, 8, 9, 9}
        Dim wcol = New Integer() {1, 2, 3, 4, 6, 7, 8, 9, 2, 8, 2, 5, 8, 2, 5, 8, 2, 4, 6, 8, 2, 4, 6, 8, 2, 4, 6, 8, 2, 4, 6, 8, 3, 7}
        Dim xrow = New Integer() {1, 1, 1, 1, 1, 1, 2, 2, 3, 3, 4, 4, 5, 6, 6, 7, 7, 8, 8, 9, 9, 9, 9, 9, 9}
        Dim xcol = New Integer() {1, 2, 3, 7, 8, 9, 2, 8, 3, 7, 4, 6, 5, 4, 6, 3, 7, 2, 8, 1, 2, 3, 7, 8, 9}
        Dim zrow = New Integer() {1, 1, 1, 1, 1, 1, 2, 2, 3, 4, 5, 6, 7, 8, 8, 9, 9, 9, 9, 9, 9}
        Dim zcol = New Integer() {1, 2, 3, 4, 5, 6, 1, 6, 5, 4, 4, 3, 2, 1, 6, 1, 2, 3, 4, 5, 6}
        Dim row1 = New Integer() {1, 1, 2, 2, 2, 3, 4, 5, 6, 7, 8, 9, 10, 10, 10, 10, 10, 10, 10}
        Dim col1 = New Integer() {3, 4, 1, 2, 4, 4, 4, 4, 4, 4, 4, 4, 1, 2, 3, 4, 5, 6, 7}
        Dim row2 = New Integer() {1, 1, 1, 1, 1, 2, 2, 3, 3, 4, 5, 6, 6, 7, 8, 9, 9, 10, 10, 10, 10, 10, 10, 10}
        Dim col2 = New Integer() {2, 3, 4, 5, 6, 1, 7, 1, 7, 7, 6, 4, 5, 3, 2, 1, 7, 1, 2, 3, 4, 5, 6, 7}
        Dim row3 = New Integer() {1, 1, 1, 1, 2, 2, 2, 3, 4, 5, 5, 5, 6, 7, 8, 9, 9, 9, 9, 9}
        Dim col3 = New Integer() {3, 4, 5, 6, 1, 2, 7, 7, 7, 4, 5, 6, 6, 7, 7, 2, 3, 4, 5, 6}
        Dim row4 = New Integer() {1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 7, 7, 7, 7, 8, 9, 10, 10, 10, 10}
        Dim col4 = New Integer() {5, 4, 5, 3, 5, 2, 5, 2, 5, 1, 5, 1, 2, 3, 4, 5, 6, 5, 5, 3, 4, 5, 6}
        Dim row5 = New Integer() {1, 1, 1, 1, 1, 2, 3, 4, 5, 5, 5, 5, 5, 6, 7, 8, 9, 9, 10, 10, 10, 10, 10}
        Dim col5 = New Integer() {2, 3, 4, 5, 6, 2, 2, 2, 2, 3, 4, 5, 6, 7, 7, 7, 1, 7, 2, 3, 4, 5, 6}
        Dim row6 = New Integer() {1, 1, 1, 1, 2, 3, 4, 5, 5, 5, 5, 5, 6, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 10, 10, 10}
        Dim col6 = New Integer() {4, 5, 6, 7, 3, 2, 1, 1, 3, 4, 5, 6, 1, 2, 7, 1, 7, 1, 7, 1, 7, 2, 3, 4, 5, 6}
        Dim row7 = New Integer() {1, 1, 1, 1, 1, 1, 1, 2, 2, 3, 4, 5, 6, 7, 8, 9, 10}
        Dim col7 = New Integer() {1, 2, 3, 4, 5, 6, 7, 2, 7, 7, 6, 6, 5, 5, 5, 4, 4}
        Dim row8 = New Integer() {1, 1, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 10, 10}
        Dim col8 = New Integer() {2, 3, 4, 5, 1, 6, 1, 6, 1, 6, 2, 3, 4, 5, 1, 6, 1, 6, 1, 6, 1, 6, 2, 3, 4, 5}
        Dim row9 = New Integer() {1, 1, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 5, 6, 6, 6, 6, 7, 8, 9, 10, 10, 10}
        Dim col9 = New Integer() {2, 3, 4, 5, 1, 6, 1, 6, 1, 6, 1, 5, 6, 2, 3, 4, 6, 6, 5, 4, 1, 2, 3}
        Dim row0 = New Integer() {1, 1, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 10, 10}
        Dim col0 = New Integer() {2, 3, 4, 5, 1, 6, 1, 6, 1, 6, 1, 6, 1, 6, 1, 6, 1, 6, 1, 6, 2, 3, 4, 5}
        Dim s1row = New Integer() {6, 6, 6, 6, 6, 6} ' minus
        Dim s1col = New Integer() {1, 2, 3, 4, 5, 6}
        Dim s2row = New Integer() {4, 4, 4, 4, 4, 4} ' equals
        Dim s2col = New Integer() {1, 2, 3, 4, 5, 6}
        Dim s3row = New Integer() {3, 4, 5, 6, 6, 6, 6, 6, 6, 6, 7, 8, 9} ' plus
        Dim s3col = New Integer() {4, 4, 4, 1, 2, 3, 4, 5, 6, 7, 4, 4, 4}
        Dim s4row = New Integer() {9, 9, 9, 9, 9, 9, 9, 9} ' underscore
        Dim s4col = New Integer() {1, 2, 3, 4, 5, 6, 7, 8}
        Dim s5row = New Integer() {3, 4, 5, 5, 5, 5, 5, 5, 5, 6, 7, 7, 8, 8}  ' asterisk
        Dim s5col = New Integer() {4, 4, 1, 2, 3, 4, 5, 6, 7, 4, 3, 5, 2, 6}
        Dim s6row = New Integer() {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12} ' open paren
        Dim s6col = New Integer() {5, 4, 4, 3, 3, 3, 3, 3, 3, 4, 4, 5}
        Dim s7row = New Integer() {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12} ' close paren
        Dim s7col = New Integer() {5, 6, 6, 7, 7, 7, 7, 7, 7, 6, 6, 5}
        Dim s8row = New Integer() {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12}  ' backslash
        Dim s8col = New Integer() {2, 2, 3, 3, 4, 4, 4, 5, 5, 6, 6, 7}
        Dim s9row = New Integer() {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12}  ' frontslash
        Dim s9col = New Integer() {7, 7, 6, 6, 5, 5, 4, 4, 3, 3, 2, 2}
        Dim s10row = New Integer() {1, 1, 1, 2, 2, 2, 3, 4, 5}  ' single quote/tick
        Dim s10col = New Integer() {4, 5, 6, 4, 5, 6, 5, 5, 5}
        Dim s11row = New Integer() {1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 4, 4}  ' double quote
        Dim s11col = New Integer() {3, 4, 6, 7, 3, 4, 6, 7, 3, 6, 3, 6}
        Dim s12row = New Integer() {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12}  ' pipe symbol
        Dim s12col = New Integer() {5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5}
        Dim s13row = New Integer() {2, 2, 3, 4, 4, 5, 6, 6, 7, 8, 8, 9, 10, 10} ' less than
        Dim s13col = New Integer() {7, 8, 6, 4, 5, 3, 1, 2, 3, 4, 5, 6, 7, 8}
        Dim s14row = New Integer() {2, 2, 3, 4, 4, 5, 6, 6, 7, 8, 8, 9, 10, 10}  ' greater than
        Dim s14col = New Integer() {1, 2, 3, 4, 5, 6, 7, 8, 6, 4, 5, 3, 1, 2}
        Dim s15row = New Integer() {1, 1, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 12, 12}  ' open bracket
        Dim s15col = New Integer() {4, 5, 6, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 5, 6}
        Dim s16row = New Integer() {1, 1, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 12, 12}  ' close bracket
        Dim s16col = New Integer() {4, 5, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 4, 5, 6}
        Dim s17row = New Integer() {2, 2, 3, 3, 6, 6, 7, 8, 8, 9}    ' semicolon
        Dim s17col = New Integer() {5, 6, 5, 6, 5, 6, 5, 4, 5, 4}
        Dim s18row = New Integer() {2, 2, 3, 3, 7, 7, 8, 8}    ' colon
        Dim s18col = New Integer() {5, 6, 5, 6, 5, 6, 5, 6}
        Dim s19row = New Integer() {9, 9, 10, 10}  ' period
        Dim s19col = New Integer() {5, 6, 5, 6}
        Dim s20row = New Integer() {1, 1, 1, 1, 2, 2, 3, 3, 4, 5, 6, 8, 8, 9, 9}  ' question mark
        Dim s20col = New Integer() {3, 4, 5, 6, 2, 7, 2, 7, 7, 6, 5, 4, 5, 4, 5}
        Dim s21row = New Integer() {2, 2, 2, 3, 4, 5, 5, 6, 6, 6, 7, 7, 8, 8, 8, 8}  ' &
        Dim s21col = New Integer() {4, 5, 6, 3, 3, 3, 4, 2, 4, 6, 2, 5, 3, 4, 5, 6}
        Dim s22row = New Integer() {8, 8, 9, 10, 10, 11}         ' comma
        Dim s22col = New Integer() {4, 5, 4, 3, 4, 3}
        Dim s23row = New Integer() {1, 2, 3, 4, 5, 6, 7, 8, 8, 8, 9, 9, 9}  ' exclamation point
        Dim s23col = New Integer() {5, 5, 5, 5, 5, 5, 5, 4, 5, 6, 4, 5, 6}
        Dim s24row = New Integer() {1, 2, 3, 3, 4, 4, 5, 5}    ' caret
        Dim s24col = New Integer() {5, 5, 4, 6, 3, 7, 2, 8}
        Dim s25row = New Integer() {1, 1, 1, 1, 2, 2, 3, 3, 4, 4, 4, 4, 5, 5, 5, 6, 6, 6, 7, 7, 7, 8, 8, 8, 8, 9, 10, 10, 11, 11, 11} ' @
        Dim s25col = New Integer() {3, 4, 5, 6, 2, 7, 1, 7, 1, 5, 6, 7, 1, 4, 7, 1, 4, 7, 1, 4, 7, 1, 5, 6, 7, 1, 2, 6, 3, 4, 5}
        Dim s26row = New Integer() {1, 1, 2, 2, 3, 3, 4, 4, 4, 4, 4, 4, 5, 5, 6, 6, 7, 7, 7, 7, 7, 7, 8, 8, 9, 9, 10, 10} ' pound sign
        Dim s26col = New Integer() {3, 6, 3, 6, 2, 5, 1, 2, 3, 4, 5, 6, 2, 5, 2, 5, 1, 2, 3, 4, 5, 6, 2, 5, 1, 4, 1, 4}
        Dim s27row = New Integer() {1, 2, 2, 2, 2, 2, 3, 3, 4, 5, 6, 6, 6, 6, 7, 8, 9, 9, 10, 10, 10, 10, 10, 11, 12}  ' dollar sign
        Dim s27col = New Integer() {4, 3, 4, 5, 6, 7, 2, 7, 2, 2, 3, 4, 5, 6, 7, 7, 2, 7, 2, 3, 4, 5, 6, 4, 4}
        Dim s28row = New Integer() {1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 5, 6, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10}  ' percent
        Dim s28col = New Integer() {2, 3, 1, 4, 1, 4, 2, 3, 4, 5, 6, 1, 2, 3, 4, 5, 3, 6, 3, 6, 4, 5}
        Dim s29row = New Integer() {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11}  ' open brace
        Dim s29col = New Integer() {5, 4, 4, 4, 4, 3, 4, 4, 4, 4, 5}
        Dim s30row = New Integer() {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11}  ' close brace
        Dim s30col = New Integer() {5, 6, 6, 6, 6, 7, 6, 6, 6, 6, 5}
        Dim s31row = New Integer() {4, 4, 5, 5, 5, 6, 6}  ' tilde
        Dim s31col = New Integer() {2, 3, 1, 4, 7, 5, 6}
        Dim s32row = New Integer() {1, 2, 3}   ' grave
        Dim s32col = New Integer() {3, 4, 5}

    End Sub

    '
    ' ----- < startBMP > -----
    '
    Private Sub startBMP()
        'Dim colix As Long
        'Dim rowix As Long
        Dim filectr As Long
        Dim tif_filename As String
        Dim i As Int32


        filectr = 1
        'Open "single_ssan.asp" For Input As #23
        Dim oFile As System.IO.File
        Dim oRead As System.IO.StreamReader
        oRead = oFile.OpenText(“single_ssan.asp”)


        While Not EOF(23)
            txtLine = 10
            txtCol = 100

            tif_filename = "tifflog" & CStr(filectr) & ".bmp"

            Dim FileStream = New FileStream("tifflog.bmp", FileMode.Create)
            FileStream.Seek(0, SeekOrigin.Begin)
            'FileStream.WriteByte(CByte(bmp(ix)))

            'Open App.Path & "\" & tif_filename For Binary As #34


            w = txtWidth.Text   ' 850  ' 34 decimal
            h = txtHeight.Text       ' 5
            Call resetbmp()
            bmp(0) = &H42
            bmp(1) = &H4D
            'Call long2DWord(79, 2)   ' bytes 2-6 file size
            Call long2DWord(62, 10)  ' bytes 10-13 data offset
            bmp(14) = &H28
            Call long2DWord(w, 18)   ' bytes 18-21 width (1-relative)
            Call long2DWord(h, 22)   ' bytes 22-25 height (1-relative)
            Call long2BackWord(1, 26)    ' bytes 26-27 planes
            Call long2BackWord(1, 28)    ' bytes 28-29 bits/pixel
            bmaprowsz = RoundDw(w) * 4
            bmapsz = bmaprowsz * h   ' DWords * bytes/DWord
            Call long2DWord(bmapsz, 34)   ' bytes 34-37 bitmap data size
            bmp(58) = &HFF
            bmp(59) = &HFF
            bmp(60) = &HFF
            bmp(61) = &H0
            ' bitmap data begins at byte 62 (&H3E)
            ' start using bsz to keep track of how many bytes you're written to the file.
            bsz = 61    ' &H41
            ' Data begins at byte 62 (0-relative, 4th row, 15th byte)
            For j = 1 To h
                ' handle full and partial DWords.
                Call bits2DWords(w)
            Next j

            Call long2DWord(bsz, 2)   ' bytes 2-6 file size

            ' Call DrawWord("BRIAN GTOQPCEUYDJKLMSVWXZH1234567890-=+_*()\/.'[|<>[];:.?&")   ' '|<>[];:.?&

            i = 1
            While Not EOF(23) And i <= 135

                oRead.ReadLine()
                'Line Input #23, InLine

                If Not EOF(23) Then
                    Call DrawWord(txtLine, txtCol, Mid(InLine, 1, 112))
                    txtCol = 100
                    txtLine = txtLine + 15
                    i = i + 1
                End If
            End While  '  While Not EOF(23) And i <= 135

            '  Do all your writing in the bmp() array before writing the array to file
            Call writeBMP()  ' zero-relative

            'Close #34
            FileStream.Close()

            filectr = filectr + 1
        End While   ' While not EOF(23)
        ' Close #23
        oRead.Close()
    End Sub







End Class

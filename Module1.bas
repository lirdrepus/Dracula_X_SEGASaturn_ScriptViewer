Attribute VB_Name = "genScript"
Option Explicit

'===========================================================
' fPrint
'   Draws an 8x8 tile on the form using CHR font data.
'
'   x, y        : Screen position to draw the tile
'   fn          : Tile index (0-based)
'   fFile       : CHR file name (without extension)
'   numForFlip  : Flip mode (8 = vertical block flip, 4 = horizontal nibble flip)
'===========================================================
Sub fPrint(x As Integer, y As Integer, fn As Byte, fFile As String, numForFlip As Byte)

    Dim Size As Integer
    Dim fStart As Long
    Dim dat(31) As Byte      ' Stores 32 bytes of tile data
    Dim tmp(31) As Byte      ' Temporary buffer for flipping
    Dim i As Integer
    Dim block As Integer
    Dim n As Integer
    Dim X1 As Integer, Y1 As Integer

    '-------------------------------------------------------
    ' Read tile data from CHR file
    '-------------------------------------------------------

    fStart = CLng("&H" & Form1.fontStart.Text)   ' Starting address of font data
    Size = Val(Form1.fSize.Text)                 ' Pixel size of each tile block

    Open fFile & ".CHR" For Binary As #1
        ' Read 32 bytes for tile fn
        Get #1, fStart + (fn * 32) + 1, dat
    Close #1


    '-------------------------------------------------------
    ' Flip mode: 8 ¡ú reverse 8 blocks of 4 bytes each
    '-------------------------------------------------------
    If numForFlip = 8 Then

        ' Copy original data
        For i = 0 To 31
            tmp(i) = dat(i)
        Next i

        ' Reverse block order (each block = 4 bytes)
        For block = 0 To 7
            For i = 0 To 3
                dat(28 - block * 4 + i) = tmp(block * 4 + i)
            Next i
        Next block


    '-------------------------------------------------------
    ' Flip mode: 4 ¡ú reverse each 4-byte block internally
    '            and swap values 1 ? 16 when flipping
    '-------------------------------------------------------
    ElseIf numForFlip = 4 Then

        ' Copy original data
        For i = 0 To 31
            tmp(i) = dat(i)
        Next i

        ' Reverse each 4-byte block
        For block = 0 To 7
            For i = 0 To 3

                ' Swap pixel values 1 ? 16 when flipping
                If tmp(block * 4 + i) = 16 Then
                    dat(block * 4 + 3 - i) = 1
                ElseIf tmp(block * 4 + i) = 1 Then
                    dat(block * 4 + 3 - i) = 16
                Else
                    dat(block * 4 + 3 - i) = tmp(block * 4 + i)
                End If

            Next i
        Next block

    End If


    '-------------------------------------------------------
    ' Draw tile on screen
    ' Each byte encodes two pixels: 01, 10, or 11
    '-------------------------------------------------------

    n = 0

    For Y1 = 0 To 7              ' 8 rows
        For X1 = 0 To 7 Step 2   ' 4 bytes per row (each byte = 2 pixels)

            Select Case True

                ' 01 ¡ú draw right pixel
                Case dat(n) = 1
                    Form1.Line _
                        ((x + (X1 + 1) * Size), (y + Y1 * Size))- _
                        ((x + (X1 + 2) * Size), (y + (Y1 + 1) * Size)), _
                        vbBlack, BF

                ' 10 ¡ú draw left pixel
                Case dat(n) = 16
                    Form1.Line _
                        ((x + X1 * Size), (y + Y1 * Size))- _
                        ((x + (X1 + 1) * Size), (y + (Y1 + 1) * Size)), _
                        vbBlack, BF

                ' 11 ¡ú draw both pixels
                Case dat(n) = 17
                    ' Left pixel
                    Form1.Line _
                        ((x + X1 * Size), (y + Y1 * Size))- _
                        ((x + (X1 + 1) * Size), (y + (Y1 + 1) * Size)), _
                        vbBlack, BF

                    ' Right pixel
                    Form1.Line _
                        ((x + (X1 + 1) * Size), (y + Y1 * Size))- _
                        ((x + (X1 + 2) * Size), (y + (Y1 + 1) * Size)), _
                        vbBlack, BF

            End Select

            n = n + 1   ' Move to next byte

        Next X1
    Next Y1

End Sub



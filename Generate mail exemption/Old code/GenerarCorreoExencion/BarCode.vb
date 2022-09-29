Imports iTextSharp.text.pdf
Imports iTextSharp.text.pdf.BarcodeCodabar
Imports iTextSharp.text.pdf.BarcodeEAN
Imports iTextSharp.text.pdf.Barcode39
Imports iTextSharp.text.pdf.Barcode128
Imports iTextSharp.text.pdf.BarcodePDF417
Imports iTextSharp.text.pdf.BarcodeDatamatrix

Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D

Public Class BarCode
    Public Function CodeCodABAR(ByVal _code As String, Optional ByVal PrintTextInCode As Boolean = False, Optional ByVal Height As Single = 0, Optional ByVal GenerateChecksum As Boolean = False, Optional ByVal ChecksumText As Boolean = False) As Bitmap
        If _code.Trim = "" Then
            Return Nothing
        Else
            Dim barcode As New BarcodeCodabar
            barcode.StartStopText = True
            barcode.GenerateChecksum = GenerateChecksum
            barcode.ChecksumText = ChecksumText
            If Height <> 0 Then barcode.BarHeight = Height
            barcode.Code = _code
            Try
                Dim bm As New System.Drawing.Bitmap(barcode.CreateDrawingImage(System.Drawing.Color.Black, System.Drawing.Color.White))
                If PrintTextInCode = False Then
                    Return bm
                Else
                    Dim bmT As Image
                    bmT = New Bitmap(bm.Width, bm.Height + 14)
                    Dim g As Graphics = Graphics.FromImage(bmT)
                    g.FillRectangle(New SolidBrush(Color.White), 0, 0, bm.Width, bm.Height + 14)

                    Dim drawFont As New Font("Arial", 8)
                    Dim drawBrush As New SolidBrush(Color.Black)

                    Dim stringSize As New SizeF
                    stringSize = g.MeasureString(_code, drawFont)
                    Dim xCenter As Single = (bm.Width - stringSize.Width) / 2
                    Dim x As Single = xCenter
                    Dim y As Single = bm.Height

                    Dim drawFormat As New StringFormat
                    drawFormat.FormatFlags = StringFormatFlags.NoWrap

                    g.DrawImage(bm, 0, 0)
                    'en el codabar no mostrar la primera letra ni la última, ya que representan estados internos del codigo
                    Dim _ncode As String = _code.Substring(1, _code.Length - 2)
                    g.DrawString(_ncode, drawFont, drawBrush, x, y, drawFormat)

                    Return bmT
                End If
            Catch ex As Exception
                Throw New Exception("Error generating EAN13 barcode. Desc:" & ex.Message)
            End Try
        End If
    End Function


    Private Function EAN13CalcChecksum(ByVal _code As String) As String
        Try
            ' Cálculo del dígito de control EAN
            Dim iSum As Integer = 0
            Dim iSumInpar As Integer = 0
            Dim iDigit As Integer = 0
            'Dim EAN As String = "590123412345"
            Dim EAN As String = _code

            EAN = EAN.PadLeft(17, "0"c)

            For i As Integer = EAN.Length To 1 Step -1
                iDigit = Convert.ToInt32(EAN.Substring(i - 1, 1))
                If i Mod 2 <> 0 Then
                    iSumInpar += iDigit
                Else
                    iSum += iDigit
                End If
            Next

            iDigit = (iSumInpar * 3) + iSum

            Dim iCheckSum As Integer = (10 - (iDigit Mod 10)) Mod 10
            Return iCheckSum.ToString
        Catch ex As Exception
            Throw New Exception("EAN13 calculation checksum error:" & ex.Message)
            Return Nothing
        End Try
    End Function

    Public Function CodeEAN13AutoGenerateChecksum(ByVal _code As String, Optional ByVal PrintTextInCode As Boolean = False, Optional ByVal Height As Single = 0, Optional ByVal GenerateChecksum As Boolean = True, Optional ByVal ChecksumText As Boolean = True) As Bitmap
        If _code.Trim = "" Then
            Return Nothing
        Else
            Dim barcode As New BarcodeEAN
            barcode.StartStopText = True
            barcode.GenerateChecksum = GenerateChecksum
            barcode.ChecksumText = ChecksumText
            If _code.Length <> 12 Then
                Throw New Exception("EAN13 code must be 12 digits lenght. Checksum value will be calculated automatically")
            End If

            If Height <> 0 Then barcode.BarHeight = Height
            _code = _code & EAN13CalcChecksum(_code)
            barcode.Code = _code
            Try
                Dim bm As New System.Drawing.Bitmap(barcode.CreateDrawingImage(System.Drawing.Color.Black, System.Drawing.Color.White))
                If PrintTextInCode = False Then
                    Return bm
                Else
                    Dim bmT As Image
                    Dim xOffset As Integer = 10
                    bmT = New Bitmap(bm.Width + xOffset, bm.Height + 14)
                    Dim g As Graphics = Graphics.FromImage(bmT)
                    g.FillRectangle(New SolidBrush(Color.White), 0, 0, bm.Width + xOffset, bm.Height + 14)

                    Dim drawFont As New Font("Arial", 8)
                    Dim drawBrush As New SolidBrush(Color.Black)

                    Dim stringSize As New SizeF
                    stringSize = g.MeasureString(_code, drawFont)
                    Dim xCenter As Single = (bm.Width - stringSize.Width) / 2
                    Dim x As Single = xCenter
                    Dim y As Single = bm.Height

                    Dim drawFormat As New StringFormat
                    drawFormat.FormatFlags = StringFormatFlags.NoWrap

                    g.DrawImage(bm, xOffset, 0)
                    'g.DrawString(_code, drawFont, drawBrush, x, y, drawFormat)

                    If xOffset < 10 Then
                        g.DrawString(_code.Substring(0, 1), drawFont, drawBrush, 0, y, drawFormat)
                    Else
                        g.DrawString(_code.Substring(0, 1), drawFont, drawBrush, xOffset - 10, y, drawFormat)
                    End If

                    Dim x1 As Single = xOffset + 4
                    g.DrawString(_code.Substring(1, 6), drawFont, drawBrush, x1, y, drawFormat)

                    Dim x2 As Single = xOffset + 50
                    g.DrawString(_code.Substring(7, 6), drawFont, drawBrush, x2, y, drawFormat)

                    g.DrawLine(Pens.Black, xOffset + 0, 0, xOffset + 0, bm.Height + 8)
                    g.DrawLine(Pens.Black, xOffset + 2, 0, xOffset + 2, bm.Height + 8)
                    g.DrawLine(Pens.Black, xOffset + 46, 0, xOffset + 46, bm.Height + 8)
                    g.DrawLine(Pens.Black, xOffset + 48, 0, xOffset + 48, bm.Height + 8)
                    g.DrawLine(Pens.Black, xOffset + 92, 0, xOffset + 92, bm.Height + 8)
                    g.DrawLine(Pens.Black, xOffset + 94, 0, xOffset + 94, bm.Height + 8)


                    Return bmT
                End If
            Catch ex As Exception
                Throw New Exception("Error generating EAN13 barcode. Desc:" & ex.Message)
            End Try
        End If
    End Function


    Public Function CodeEAN13(ByVal _code As String, Optional ByVal PrintTextInCode As Boolean = False, Optional ByVal Height As Single = 0, Optional ByVal GenerateChecksum As Boolean = True, Optional ByVal ChecksumText As Boolean = True) As Bitmap
        If _code.Trim = "" Then
            Return Nothing
        Else
            Dim barcode As New BarcodeEAN
            'barcode.CodeType = iTextSharp.text.pdf.Barcode.UPCA
            barcode.StartStopText = True
            barcode.GenerateChecksum = GenerateChecksum
            barcode.ChecksumText = ChecksumText
            If _code.Length <> 13 Then
                Throw New Exception("EAN13 code must be 13 digits lenght")
            End If

            If Height <> 0 Then barcode.BarHeight = Height
            barcode.Code = _code
            Try
                Dim bm As New System.Drawing.Bitmap(barcode.CreateDrawingImage(System.Drawing.Color.Black, System.Drawing.Color.White))
                If PrintTextInCode = False Then
                    Return bm
                Else
                    Dim bmT As Image
                    Dim xOffset As Integer = 10
                    bmT = New Bitmap(bm.Width + xOffset, bm.Height + 14)
                    Dim g As Graphics = Graphics.FromImage(bmT)
                    g.FillRectangle(New SolidBrush(Color.White), 0, 0, bm.Width + xOffset, bm.Height + 14)

                    Dim drawFont As New Font("Arial", 8)
                    Dim drawBrush As New SolidBrush(Color.Black)

                    Dim stringSize As New SizeF
                    stringSize = g.MeasureString(_code, drawFont)
                    Dim xCenter As Single = (bm.Width - stringSize.Width) / 2
                    Dim x As Single = xCenter
                    Dim y As Single = bm.Height

                    Dim drawFormat As New StringFormat
                    drawFormat.FormatFlags = StringFormatFlags.NoWrap

                    g.DrawImage(bm, xOffset, 0)
                    'g.DrawString(_code, drawFont, drawBrush, x, y, drawFormat)

                    If xOffset < 10 Then
                        g.DrawString(_code.Substring(0, 1), drawFont, drawBrush, 0, y, drawFormat)
                    Else
                        g.DrawString(_code.Substring(0, 1), drawFont, drawBrush, xOffset - 10, y, drawFormat)
                    End If

                    Dim x1 As Single = xOffset + 4
                    g.DrawString(_code.Substring(1, 6), drawFont, drawBrush, x1, y, drawFormat)

                    Dim x2 As Single = xOffset + 50
                    g.DrawString(_code.Substring(7, 6), drawFont, drawBrush, x2, y, drawFormat)

                    g.DrawLine(Pens.Black, xOffset + 0, 0, xOffset + 0, bm.Height + 8)
                    g.DrawLine(Pens.Black, xOffset + 2, 0, xOffset + 2, bm.Height + 8)
                    g.DrawLine(Pens.Black, xOffset + 46, 0, xOffset + 46, bm.Height + 8)
                    g.DrawLine(Pens.Black, xOffset + 48, 0, xOffset + 48, bm.Height + 8)
                    g.DrawLine(Pens.Black, xOffset + 92, 0, xOffset + 92, bm.Height + 8)
                    g.DrawLine(Pens.Black, xOffset + 94, 0, xOffset + 94, bm.Height + 8)


                    Return bmT
                End If
            Catch ex As Exception
                Throw New Exception("Error generating EAN13 barcode. Desc:" & ex.Message)
            End Try
        End If
    End Function

    Public Enum Code128SubTypes
        'CODABAR = iTextSharp.text.pdf.Barcode.CODABAR
        CODE128 = iTextSharp.text.pdf.Barcode.CODE128
        CODE128_RAW = iTextSharp.text.pdf.Barcode.CODE128_RAW
        CODE128_UCC = iTextSharp.text.pdf.Barcode.CODE128_UCC
        'EAN13 = iTextSharp.text.pdf.Barcode.EAN13
        'EAN8 = iTextSharp.text.pdf.Barcode.EAN8
        'PLANET = iTextSharp.text.pdf.Barcode.PLANET
        'POSTNET = iTextSharp.text.pdf.Barcode.POSTNET
        'SUPP2 = iTextSharp.text.pdf.Barcode.SUPP2
        'SUPP5 = iTextSharp.text.pdf.Barcode.SUPP5
        'UPCA = iTextSharp.text.pdf.Barcode.UPCA
        'UPCE = iTextSharp.text.pdf.Barcode.UPCE
    End Enum
    Public Function Code128(ByVal _code As String, Optional ByVal codeType As Integer = Code128SubTypes.CODE128, Optional ByVal PrintTextInCode As Boolean = False, Optional ByVal Height As Single = 0, Optional ByVal GenerateChecksum As Boolean = True, Optional ByVal ChecksumText As Boolean = True) As Bitmap
        If _code.Trim = "" Then
            Return Nothing
        Else
            Dim barcode As New Barcode128

            barcode.CodeType = codeType
            barcode.StartStopText = True
            barcode.GenerateChecksum = GenerateChecksum
            barcode.ChecksumText = ChecksumText
            If Height <> 0 Then barcode.BarHeight = Height
            barcode.Code = _code
            Try
                Dim bm As New System.Drawing.Bitmap(barcode.CreateDrawingImage(System.Drawing.Color.Black, System.Drawing.Color.White))
                If PrintTextInCode = False Then
                    Return bm
                Else
                    Dim bmT As Image
                    bmT = New Bitmap(bm.Width, bm.Height + 14)
                    Dim g As Graphics = Graphics.FromImage(bmT)
                    g.FillRectangle(New SolidBrush(Color.White), 0, 0, bm.Width, bm.Height + 14)

                    Dim drawFont As New Font("Arial", 8)
                    Dim drawBrush As New SolidBrush(Color.Black)

                    Dim stringSize As New SizeF
                    stringSize = g.MeasureString(_code, drawFont)
                    Dim xCenter As Single = (bm.Width - stringSize.Width) / 2
                    Dim x As Single = xCenter
                    Dim y As Single = bm.Height

                    Dim drawFormat As New StringFormat
                    drawFormat.FormatFlags = StringFormatFlags.NoWrap

                    g.DrawImage(bm, 0, 0)
                    g.DrawString(_code, drawFont, drawBrush, x, y, drawFormat)

                    Return bmT
                End If
            Catch ex As Exception
                Throw New Exception("Error generating code128 barcode. Desc:" & ex.Message)
            End Try
        End If
    End Function


    Public Function Code39(ByVal _code As String, Optional ByVal PrintTextInCode As Boolean = False, Optional ByVal Height As Single = 0, Optional ByVal GenerateChecksum As Boolean = False, Optional ByVal ChecksumText As Boolean = False) As Bitmap
        If _code.Trim = "" Then
            Return Nothing
        Else
            Dim barcode As New Barcode39
            barcode.StartStopText = True
            barcode.GenerateChecksum = GenerateChecksum
            barcode.ChecksumText = ChecksumText

            If Height <> 0 Then barcode.BarHeight = Height
            barcode.Code = _code
            Try
                Dim bm As New System.Drawing.Bitmap(barcode.CreateDrawingImage(System.Drawing.Color.Black, System.Drawing.Color.White))
                If PrintTextInCode = False Then
                    Return bm
                Else
                    Dim bmT As Image
                    bmT = New Bitmap(bm.Width, bm.Height + 14)
                    Dim g As Graphics = Graphics.FromImage(bmT)
                    g.FillRectangle(New SolidBrush(Color.White), 0, 0, bm.Width, bm.Height + 14)

                    Dim drawFont As New Font("Arial", 8)
                    Dim drawBrush As New SolidBrush(Color.Black)

                    Dim stringSize As New SizeF
                    stringSize = g.MeasureString(_code, drawFont)
                    Dim xCenter As Single = (bm.Width - stringSize.Width) / 2
                    Dim x As Single = xCenter
                    Dim y As Single = bm.Height

                    Dim drawFormat As New StringFormat
                    drawFormat.FormatFlags = StringFormatFlags.NoWrap

                    g.DrawImage(bm, 0, 0)
                    g.DrawString(_code, drawFont, drawBrush, x, y, drawFormat)

                    Return bmT
                End If

            Catch ex As Exception
                Throw New Exception("Error generating code39 barcode. Desc:" & ex.Message)
            End Try
        End If
    End Function

    Public Function DataMatrix(ByVal _code As String, Optional ByVal Scale As Integer = 1) As Bitmap
        If _code.Trim = "" Then
            Return Nothing
        Else
            Dim barcode As BarcodeDatamatrix = New BarcodeDatamatrix()
            Dim barcodeDimensions() As Integer = New Integer() {10, 12, 14, 16, 18, 20, 22, 24, 26, 32, 36, 40, 44, 48, 52, 64, 72, 80, 88, 96, 104, 120, 132, 144}
            Dim returnResult As Integer
            Try
                Dim bm As System.Drawing.Bitmap = Nothing
                'barcode.Options = BarcodeDatamatrix.DM_AUTO
                barcode.Options = BarcodeDatamatrix.DM_B256
                For generateCount As Integer = 0 To barcodeDimensions.Length - 1
                    barcode.Width = barcodeDimensions(generateCount)
                    barcode.Height = barcodeDimensions(generateCount)
                    returnResult = barcode.Generate(_code)
                    If returnResult = BarcodeDatamatrix.DM_NO_ERROR Then
                        bm = New System.Drawing.Bitmap(barcode.CreateDrawingImage(System.Drawing.Color.Black, System.Drawing.Color.White))
                        If Scale <> 1 Then
                            'Dim original As Image = bm
                            Dim finalW As Integer = CType(bm.Width * Scale, Integer)
                            Dim finalH As Integer = CType(bm.Height * Scale, Integer)

                            Dim retBitmap As New Bitmap(finalW, finalH)
                            Dim retgr As Graphics = Graphics.FromImage(retBitmap)
                            retgr.ScaleTransform(Scale, Scale)
                            retgr.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
                            retgr.InterpolationMode = Drawing2D.InterpolationMode.NearestNeighbor
                            retgr.PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality

                            retgr.DrawImage(bm, New Point(0, 0))

                            Return retBitmap
                        End If
                        Exit For
                    End If
                Next
                Return bm
            Catch ex As Exception
                Throw New Exception("Error generating datamatrix barcode. Desc:" & ex.Message)
            End Try
        End If
    End Function

    Public Function PDF417(ByVal _code As String, Optional ByVal Scale As Integer = 1) As Bitmap
        If _code.Trim = "" Then
            Return Nothing
        Else
            Dim barcode As New BarcodePDF417
            barcode.Options = BarcodePDF417.PDF417_USE_ASPECT_RATIO
            'barcode.YHeight = 6
            barcode.ErrorLevel = 8
            barcode.SetText(_code)

            'Dim encoding As New System.Text.UTF8Encoding
            'Dim b() As Byte = encoding.GetBytes(_code)
            'barcode.Text = b
            Try
                Dim bm As New System.Drawing.Bitmap(barcode.CreateDrawingImage(System.Drawing.Color.Black, System.Drawing.Color.White))
                'Return bm
                If Scale <> 1 Then
                    'Dim original As Image = bm
                    Dim finalW As Integer = CType(bm.Width * Scale, Integer)
                    Dim finalH As Integer = CType(bm.Height * Scale, Integer)

                    Dim retBitmap As New Bitmap(finalW, finalH)
                    Dim retgr As Graphics = Graphics.FromImage(retBitmap)
                    retgr.ScaleTransform(Scale, Scale)
                    retgr.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
                    retgr.InterpolationMode = Drawing2D.InterpolationMode.NearestNeighbor
                    retgr.PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality

                    retgr.DrawImage(bm, New Point(0, 0))

                    Return retBitmap
                Else
                    Return bm
                End If
            Catch ex As Exception
                Throw New Exception("Error generating PDF417 barcode. Desc:" & ex.Message)
            End Try
        End If
    End Function

End Class

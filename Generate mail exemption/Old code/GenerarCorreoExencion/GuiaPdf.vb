Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System
Imports System.Globalization
Imports System.IO

Public Class RoundedBorder : Implements IPdfPCellEvent
    Public Sub CellLayout(cell As PdfPCell, rect As Rectangle, canvas As PdfContentByte()) Implements IPdfPCellEvent.CellLayout
        Dim cb As PdfContentByte = canvas(PdfPTable.BACKGROUNDCANVAS)
        cb.RoundRectangle(rect.Left + 1.5F, rect.Bottom + 1.5F, rect.Width - 3, rect.Height - 3, 4)
        cb.Stroke()
    End Sub
End Class

Public Class GuiaPdf
    Dim mostrar2 As New Cargar
    Dim enviar As New Envio_De_Correos
    Dim CodigoBarras As New BarCode
    Dim MyconString As String
    Dim Consulta As String

    Dim imagepath As String = "C:\inetpub\wwwroot\images" 'Server.MapPath("Images") 'path donde están las imágenes

    Dim Parrafo As Paragraph
    Dim Frase As Phrase
    Dim SubParrafo1, SubParrafo2, SubParrafo3, SubParrafo4, SubParrafo5, SubParrafo6 As Chunk
    Dim Celda As PdfPCell
    Dim Enlace As Anchor

    'crear fuente y tamaño de la fuente
    Dim Arial5 As Font = FontFactory.GetFont("Arial", 5, iTextSharp.text.Font.NORMAL, BaseColor.BLACK)

    Dim Arial9 As Font = FontFactory.GetFont("Arial", 9, iTextSharp.text.Font.NORMAL, BaseColor.BLACK)
    Dim Arial9Negrita As Font = FontFactory.GetFont("Arial", 9, iTextSharp.text.Font.BOLD, BaseColor.BLACK)
    Dim Arial10 As Font = FontFactory.GetFont("Arial", 10, iTextSharp.text.Font.NORMAL, BaseColor.BLACK)
    Dim Arial10Negrita As Font = FontFactory.GetFont("Arial", 10, iTextSharp.text.Font.BOLD, BaseColor.BLACK)
    Dim Arial10NegritaSubrayado As Font = FontFactory.GetFont("Arial", 12, iTextSharp.text.Font.BOLD, BaseColor.BLACK)
    Dim Arial12 As Font = FontFactory.GetFont("Arial", 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK)
    Dim Arial12Negrita As Font = FontFactory.GetFont("Arial", 12, iTextSharp.text.Font.BOLD, BaseColor.BLACK)
    Dim Arial14 As Font = FontFactory.GetFont("Arial", 14, iTextSharp.text.Font.NORMAL, BaseColor.BLACK)
    Dim Arial14Negrita As Font = FontFactory.GetFont("Arial", 14, iTextSharp.text.Font.BOLD, BaseColor.BLACK)
    Dim Arial16 As Font = FontFactory.GetFont("Arial", 16, iTextSharp.text.Font.NORMAL, BaseColor.BLACK)
    Dim Arial16Negrita As Font = FontFactory.GetFont("Arial", 16, iTextSharp.text.Font.BOLD, BaseColor.BLACK)


    '>> Gloria Tarea 234 21-feb-2018
    Dim CodigoImportador As Integer = 0
    Dim NombreImportador As String = ""
    Dim Linea1AWB As String = ""
    Dim Linea2AWB As String = ""
    Dim Linea3AWB As String = ""

    Dim CodigoPaquete As Integer
    Dim EnviadoPor, Descripcion, Tracking As String
    Dim Peso As Decimal

    Function Crear_Celda(ByVal Texto As String, ByVal fuente As iTextSharp.text.Font, ByVal alineacion As Integer) As PdfPCell
        Dim Frase As Phrase
        Dim SubParrafo1 As Chunk

        Frase = New Phrase()

        SubParrafo1 = New Chunk(Texto, fuente)
        Frase.Add(SubParrafo1)
        Celda = New PdfPCell(Frase)
        Celda.HorizontalAlignment = alineacion
        Celda.Border = Rectangle.NO_BORDER

        Crear_Celda = Celda
    End Function

    Function Crear_Celda_Padding(ByVal Texto As String, ByVal fuente As iTextSharp.text.Font, ByVal alineacion As Integer) As PdfPCell
        Dim Frase As Phrase
        Dim SubParrafo1 As Chunk

        Frase = New Phrase()

        SubParrafo1 = New Chunk(Texto, fuente)
        Frase.Add(SubParrafo1)
        Celda = New PdfPCell(Frase)
        Celda.HorizontalAlignment = alineacion
        Celda.Border = Rectangle.NO_BORDER
        Celda.PaddingBottom = 8

        Crear_Celda_Padding = Celda
    End Function

    Function Crear_Celda_Padding7(ByVal Texto As String, ByVal fuente As iTextSharp.text.Font, ByVal alineacion As Integer) As PdfPCell
        Dim Frase As Phrase
        Dim SubParrafo1 As Chunk

        Frase = New Phrase()

        SubParrafo1 = New Chunk(Texto, fuente)
        Frase.Add(SubParrafo1)
        Celda = New PdfPCell(Frase)
        Celda.HorizontalAlignment = alineacion
        Celda.VerticalAlignment = Element.ALIGN_TOP
        Celda.Border = Rectangle.NO_BORDER
        Celda.PaddingBottom = 7 '8

        Crear_Celda_Padding7 = Celda
    End Function

    'Gloria Tarea 245 23-feb-2018, Separar los formatos de exención por Importador
    'Just here
    Function Crear_Reporte_Pdf(
        ByVal Guia As String, 
        ByRef NombreArchivo As String, 
        ByVal conString As String, 
        ByRef Mensaje As String
        ) As Boolean
        Dim bResultado As Boolean = False
        Dim EnviarExencion As Boolean

        Try
            MyconString = conString

            Mensaje = ""
            EnviadoPor = "" : Descripcion = "" : Tracking = ""
            Obtener_Datos_Guia(Guia, CodigoPaquete, EnviadoPor, Descripcion, Peso, Tracking)
            Obtener_Datos_Importador(Guia, EnviarExencion) '>> Gloria Tarea 234 21-feb-2018

            Select Case CodigoImportador
                Case 4 'CPX
                    bResultado = Crear_Reporte_Pdf_CPX(Guia, NombreArchivo, Mensaje)
                Case 1, 6, 7, 8, 9 'PidoBox RapiPuerta QS ICC YoCargo
                    bResultado = Crear_Reporte_Pdf_PidoBox(Guia, NombreArchivo, Mensaje, CodigoImportador)
                Case Else 'no hay importador
                    If EnviarExencion = True Then 'verifica si la línea de pedido se debe enviar correo, de ser así mandar alerta de error
                        Mensaje = "No se encontró el código de Importador en el pedido. Guía : " & Guia
                    Else 'en el campo observaciones se agregó No-Enviar_Exencion y el proceso la debe ignorar
                        bResultado = True
                        NombreArchivo = ""
                    End If
            End Select
        Catch ex As Exception
            Mensaje = ex.Message.ToString
            bResultado = False
        End Try

        Return bResultado
    End Function

    '***********************************************************
    ' CREAR PDF DE CPX
    '***********************************************************

    Function Crear_Reporte_Pdf_CPX(ByVal Guia As String, ByRef NombreArchivo As String, ByRef Mensaje As String) As Boolean
        Try
            Dim TablaGeneral As PdfPTable = New PdfPTable(1)
            TablaGeneral.DefaultCell.Border = Rectangle.NO_BORDER
            TablaGeneral.WidthPercentage = 100

            Celda = New PdfPCell(Obtener_Encabezado_CPX(Guia))
            Celda.HorizontalAlignment = Element.ALIGN_LEFT
            Celda.Border = Rectangle.NO_BORDER
            TablaGeneral.AddCell(Celda)

            Celda = New PdfPCell(Obtener_Direccion_Empresa_CPX(EnviadoPor))
            Celda.HorizontalAlignment = Element.ALIGN_CENTER
            Celda.Border = Rectangle.NO_BORDER
            TablaGeneral.AddCell(Celda)

            Frase = New Phrase()
            SubParrafo1 = New Chunk(vbNewLine, Arial12)
            Frase.Add(SubParrafo1)
            TablaGeneral.AddCell(Frase)

            Celda = New PdfPCell(Obtener_Descripcion_CPX(Descripcion, Peso, Tracking))
            Celda.HorizontalAlignment = Element.ALIGN_LEFT
            Celda.Border = Rectangle.NO_BORDER
            TablaGeneral.AddCell(Celda)

            Frase = New Phrase()
            SubParrafo1 = New Chunk("THE GOODS HAVE BEEN RECEIVED IN APPARENTLY GOOD CONDITIONS WE" + vbNewLine + "ARE NOT RESPONSIBLES FOR CONCEALED DAMAGES OR SHORTAGES", Arial12)
            Frase.Add(SubParrafo1)
            Celda = New PdfPCell(Frase)
            Celda.HorizontalAlignment = Element.ALIGN_LEFT
            Celda.Border = Rectangle.NO_BORDER
            Celda.PaddingTop = 95
            TablaGeneral.AddCell(Celda)

            Celda = New PdfPCell(Obtener_Pie_CPX(Tracking))
            Celda.HorizontalAlignment = Element.ALIGN_CENTER
            Celda.Border = Rectangle.NO_BORDER
            Celda.PaddingTop = 50
            TablaGeneral.AddCell(Celda)


            'documento tamaño carta
            Dim doc1 = New Document(iTextSharp.text.PageSize.LETTER)
            Dim writer As PdfWriter = PdfWriter.GetInstance(doc1, New FileStream("C:\inetpub\wwwroot\Sistema\CARGA\" + Tracking + ".pdf", FileMode.Create))
            NombreArchivo = "C:\inetpub\wwwroot\Sistema\CARGA\" + Tracking + ".pdf"
            doc1.Open()

            doc1.Add(TablaGeneral)

            doc1.Close()
            writer.Close()

            Return True
        Catch ex As Exception
            Mensaje = ex.Message.ToString
            Return False
        End Try
    End Function

    Function Obtener_Encabezado_CPX(ByVal Guia As String) As PdfPTable
        Dim TablaEncabezado As PdfPTable = New PdfPTable(3)
        Dim AnchoEncabezado() As Single = {35, 40, 25}

        TablaEncabezado.WidthPercentage = 100
        TablaEncabezado.SetWidths(AnchoEncabezado)
        TablaEncabezado.DefaultCell.Border = Rectangle.NO_BORDER


        Frase = New Phrase()
        SubParrafo1 = New Chunk(NombreImportador + vbNewLine, Arial10Negrita)
        SubParrafo2 = New Chunk(Linea1AWB + vbNewLine, Arial10)
        SubParrafo3 = New Chunk(Linea2AWB + vbNewLine, Arial10)
        SubParrafo4 = New Chunk(Linea3AWB + vbNewLine, Arial10)

        Frase.Add(SubParrafo1)
        Frase.Add(SubParrafo2)
        Frase.Add(SubParrafo3)
        Frase.Add(SubParrafo4)

        Celda = New PdfPCell(Frase)
        Celda.HorizontalAlignment = Element.ALIGN_LEFT
        Celda.Border = Rectangle.NO_BORDER
        Celda.BorderWidthTop = 1
        Celda.BorderWidthBottom = 1
        Celda.PaddingBottom = 6
        Celda.PaddingTop = 3
        TablaEncabezado.AddCell(Celda)

        'Celda 2
        Frase = New Phrase()
        SubParrafo1 = New Chunk("AWB ", Arial16)
        SubParrafo2 = New Chunk(" " + Guia, Arial16Negrita)
        Frase.Add(SubParrafo1)
        Frase.Add(SubParrafo2)

        Celda = New PdfPCell(Frase)
        Celda.HorizontalAlignment = Element.ALIGN_LEFT
        Celda.VerticalAlignment = Element.ALIGN_BOTTOM
        Celda.Border = Rectangle.NO_BORDER
        Celda.BorderWidthTop = 1
        Celda.BorderWidthBottom = 1
        Celda.PaddingBottom = 10
        Celda.PaddingTop = 3
        TablaEncabezado.AddCell(Celda)

        'Celda 3
        'Dim bm As System.Drawing.Bitmap = CodigoBarras.CodeCodABAR("A0" + Guia + "B", False, 12)
        Dim bm As System.Drawing.Bitmap = CodigoBarras.Code128(Guia, 9, False, 13, , False) 'Números y letras

        Dim ImagenBarras As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(bm, System.Drawing.Imaging.ImageFormat.Png)


        Celda = New PdfPCell(ImagenBarras)
        Celda.HorizontalAlignment = Element.ALIGN_RIGHT
        Celda.VerticalAlignment = Element.ALIGN_TOP
        Celda.Border = Rectangle.NO_BORDER
        Celda.BorderWidthTop = 1
        Celda.BorderWidthBottom = 1
        Celda.PaddingBottom = 6
        Celda.PaddingTop = 6
        Celda.PaddingRight = 10
        TablaEncabezado.AddCell(Celda)

        Obtener_Encabezado_CPX = TablaEncabezado
    End Function

    Function Obtener_Direccion_Empresa_CPX(ByVal EnviadoPor As String) As PdfPTable
        Dim TablaEmpresas As PdfPTable = New PdfPTable(2)
        Dim AnchoEmpresas() As Single = {47, 53}
        TablaEmpresas.WidthPercentage = 100
        TablaEmpresas.SetWidths(AnchoEmpresas)
        TablaEmpresas.DefaultCell.Border = Rectangle.NO_BORDER

        Dim TablaProveedor As PdfPTable = New PdfPTable(2)
        Dim AnchoProveedor() As Single = {22, 78}
        TablaProveedor.WidthPercentage = 100
        TablaProveedor.SetWidths(AnchoProveedor)
        TablaProveedor.DefaultCell.Border = Rectangle.NO_BORDER


        Dim TablaCliente As PdfPTable = New PdfPTable(2)
        Dim AnchoCliente() As Single = {22, 78}
        TablaCliente.WidthPercentage = 100
        TablaCliente.SetWidths(AnchoCliente)
        TablaCliente.DefaultCell.Border = Rectangle.NO_BORDER

        Dim TablaRedondeada1 As PdfPTable = New PdfPTable(1)
        TablaRedondeada1.DefaultCell.Border = PdfPCell.NO_BORDER
        TablaRedondeada1.DefaultCell.CellEvent = New RoundedBorder()

        Dim TablaRedondeada2 As PdfPTable = New PdfPTable(1)
        TablaRedondeada2.DefaultCell.Border = PdfPCell.NO_BORDER
        TablaRedondeada2.DefaultCell.CellEvent = New RoundedBorder()


        TablaProveedor.AddCell(Crear_Celda("Shipper", Arial12Negrita, Element.ALIGN_LEFT))
        TablaProveedor.AddCell(Crear_Celda(EnviadoPor, Arial12, Element.ALIGN_LEFT))

        TablaProveedor.AddCell(" ")
        TablaProveedor.AddCell(" ")

        TablaProveedor.AddCell(Crear_Celda("Address", Arial12Negrita, Element.ALIGN_LEFT))
        TablaProveedor.AddCell("")

        TablaProveedor.AddCell("")
        TablaProveedor.AddCell(Crear_Celda("40", Arial12, Element.ALIGN_LEFT))

        TablaProveedor.AddCell(" ")
        TablaProveedor.AddCell(" ")


        TablaCliente.AddCell(Crear_Celda("Customer", Arial12Negrita, Element.ALIGN_LEFT))
        TablaCliente.AddCell(Crear_Celda("GUATEMALA DIGITAL S.A.", Arial12, Element.ALIGN_LEFT))

        TablaCliente.AddCell(" ")
        TablaCliente.AddCell(" ")
        TablaCliente.AddCell(" ")
        TablaCliente.AddCell(" ")


        TablaCliente.AddCell(Crear_Celda("Address", Arial12Negrita, Element.ALIGN_LEFT))
        TablaCliente.AddCell(Crear_Celda("CIUDAD" + vbNewLine + "GUATEMALA CITY GUATEMALA" + vbNewLine + "GUATEMALA", Arial12, Element.ALIGN_LEFT))

        TablaCliente.AddCell(" ")
        TablaCliente.AddCell(" ")


        TablaRedondeada1.AddCell(TablaProveedor)
        TablaRedondeada2.AddCell(TablaCliente)

        TablaEmpresas.AddCell(TablaRedondeada1)
        TablaEmpresas.AddCell(TablaRedondeada2)

        Obtener_Direccion_Empresa_CPX = TablaEmpresas

    End Function

    Function Obtener_Descripcion_CPX(ByVal Descripcion As String, ByVal Peso As Decimal, ByVal Tracking As String) As PdfPTable

        Dim Fecha As Date
        Dim CadPeso, CadFreight As String
        Dim TablaDescripcion As PdfPTable = New PdfPTable(2)
        Dim AnchoDescripcion() As Single = {15, 85}
        TablaDescripcion.WidthPercentage = 100
        TablaDescripcion.SetWidths(AnchoDescripcion)
        TablaDescripcion.DefaultCell.Border = Rectangle.NO_BORDER

        Dim TrackingReducido, MensajeError, CadPiezas As String

        If Peso > 0 Then
            CadPeso = Format(Peso, "##,##0") & " Pounds"
            CadFreight = "$" + Format(Peso * 1.6, "##,##0.00")
        Else
            CadPeso = "NULL"
            CadFreight = "NULL"
        End If


        TrackingReducido = "" : MensajeError = "" : CadPiezas = "null"
        If Validar_Tracking(Tracking, TrackingReducido, MensajeError) = True Then

            If TrackingReducido <> "" Then
                Consulta = "select SUM(cantidad) from pedido where codigoderastreo like '%" & TrackingReducido & "%' and CodigoEstadoPedido <> 4 "
                CadPiezas = CStr(mostrar2.retornarentero(Consulta, MyconString))

                Consulta = "select top(1) FechaDeEnvio from pedido where codigoderastreo like '%" & TrackingReducido & "%'  and CodigoEstadoPedido <> 4 order by FechaDeEnvio desc"
                Fecha = Obtener_Fecha_envio(Consulta)

            Else
                Consulta = "select SUM(cantidad) from pedido where codigoderastreo like '%" & Tracking & "%' and CodigoEstadoPedido <> 4 "
                CadPiezas = CStr(mostrar2.retornarentero(Consulta, MyconString))

                Consulta = "select top(1) FechaDeEnvio from pedido where codigoderastreo like '%" & Tracking & "%'  and CodigoEstadoPedido <> 4 order by FechaDeEnvio desc"
                Fecha = Obtener_Fecha_envio(Consulta)

            End If

            If CadPiezas = "0" Then
                CadPiezas = "null"
            End If


            If DateAdd(DateInterval.Day, 20, Fecha) > Date.Now Then
                Fecha = Date.Now
            Else
                Fecha = DateAdd(DateInterval.Day, 20, Fecha)
            End If

        End If

        TablaDescripcion.AddCell(Crear_Celda_Padding("Date", Arial12Negrita, Element.ALIGN_RIGHT))
        TablaDescripcion.AddCell(Crear_Celda_Padding(Obtener_Fecha(Fecha), Arial12, Element.ALIGN_LEFT))

        TablaDescripcion.AddCell(Crear_Celda_Padding("Origen", Arial12Negrita, Element.ALIGN_RIGHT))
        TablaDescripcion.AddCell(Crear_Celda_Padding("", Arial12, Element.ALIGN_LEFT))

        TablaDescripcion.AddCell(Crear_Celda_Padding("Flight No.", Arial12Negrita, Element.ALIGN_RIGHT))
        TablaDescripcion.AddCell(Crear_Celda_Padding("null", Arial12, Element.ALIGN_LEFT))

        TablaDescripcion.AddCell(Crear_Celda_Padding("Destination", Arial12Negrita, Element.ALIGN_RIGHT))
        TablaDescripcion.AddCell(Crear_Celda_Padding("GUA", Arial12, Element.ALIGN_LEFT))

        TablaDescripcion.AddCell(Crear_Celda_Padding("Description", Arial12Negrita, Element.ALIGN_RIGHT))
        TablaDescripcion.AddCell(Crear_Celda_Padding(Descripcion, Arial12, Element.ALIGN_LEFT))

        TablaDescripcion.AddCell(Crear_Celda_Padding("Pieces", Arial12Negrita, Element.ALIGN_RIGHT))
        TablaDescripcion.AddCell(Crear_Celda_Padding(CadPiezas, Arial12, Element.ALIGN_LEFT))

        TablaDescripcion.AddCell(Crear_Celda_Padding("Weight", Arial12Negrita, Element.ALIGN_RIGHT))
        TablaDescripcion.AddCell(Crear_Celda_Padding(CadPeso, Arial12, Element.ALIGN_LEFT))
        'TablaDescripcion.AddCell(Crear_Celda_Padding(Format(Peso, "##,##0") & " Pounds", Arial12, Element.ALIGN_LEFT))

        TablaDescripcion.AddCell(Crear_Celda_Padding("Tracking", Arial12Negrita, Element.ALIGN_RIGHT))
        TablaDescripcion.AddCell(Crear_Celda_Padding(Tracking, Arial12, Element.ALIGN_LEFT))

        TablaDescripcion.AddCell(Crear_Celda_Padding("Freight", Arial12Negrita, Element.ALIGN_RIGHT))
        TablaDescripcion.AddCell(Crear_Celda_Padding(CadFreight, Arial12, Element.ALIGN_LEFT))
        'TablaDescripcion.AddCell(Crear_Celda_Padding("$" + Format(Peso * 1.6, "##,##0.00"), Arial12, Element.ALIGN_LEFT))

        Obtener_Descripcion_CPX = TablaDescripcion

    End Function

    Function Obtener_Pie_CPX(ByVal Tracking As String) As PdfPTable
        Dim Fecha As Date
        Dim TrackingReducido, MensajeError As String

        Dim TablaPie As PdfPTable = New PdfPTable(3)
        Dim AnchoPie() As Single = {30, 42, 28}
        TablaPie.WidthPercentage = 100
        TablaPie.SetWidths(AnchoPie)
        TablaPie.DefaultCell.Border = Rectangle.NO_BORDER

        TrackingReducido = "" : MensajeError = ""
        If Validar_Tracking(Tracking, TrackingReducido, MensajeError) = True Then

            If TrackingReducido <> "" Then
                Consulta = "select top(1) FechaDeEnvio from pedido where codigoderastreo like '%" & TrackingReducido & "%' order by FechaDeEnvio desc"
                Fecha = Obtener_Fecha_envio(Consulta)
            Else
                Consulta = "select top(1) FechaDeEnvio from pedido where codigoderastreo like '%" & Tracking & "%' order by FechaDeEnvio desc"
                Fecha = Obtener_Fecha_envio(Consulta)
            End If

        End If

        If DateAdd(DateInterval.Day, 20, Fecha) > Date.Now Then
            Fecha = Date.Now
        Else
            Fecha = DateAdd(DateInterval.Day, 20, Fecha)
        End If

        Frase = New Phrase()
        SubParrafo1 = New Chunk("" + vbNewLine, Arial10Negrita)
        SubParrafo2 = New Chunk("___________________________" + vbNewLine, Arial10)
        SubParrafo3 = New Chunk("RECEIVED BY", Arial10)
        Frase.Add(SubParrafo1)
        Frase.Add(SubParrafo2)
        Frase.Add(SubParrafo3)
        Celda = New PdfPCell(Frase)
        Celda.HorizontalAlignment = Element.ALIGN_CENTER
        Celda.Border = Rectangle.NO_BORDER
        TablaPie.AddCell(Celda)

        Frase = New Phrase()
        SubParrafo1 = New Chunk("MIAMI, FLORIDA " + Obtener_Fecha(Fecha) + "  " + vbNewLine, Arial10Negrita)
        SubParrafo2 = New Chunk("______________________________" + vbNewLine, Arial10)
        SubParrafo3 = New Chunk("DATE", Arial10)
        Frase.Add(SubParrafo1)
        Frase.Add(SubParrafo2)
        Frase.Add(SubParrafo3)
        Celda = New PdfPCell(Frase)
        Celda.HorizontalAlignment = Element.ALIGN_CENTER
        Celda.Border = Rectangle.NO_BORDER
        TablaPie.AddCell(Celda)

        TablaPie.AddCell("")

        'TablaPie.AddCell("")
        'TablaPie.AddCell(Crear_Celda("MIAMI, FLORIDA 09/14/2015", Arial10Negrita, Element.ALIGN_CENTER))


        'TablaPie.AddCell(Crear_Celda("______________________________" + vbNewLine + "RECEIVED BY", Arial10, Element.ALIGN_CENTER))
        'TablaPie.AddCell(Crear_Celda("______________________________" + vbNewLine + "DATE", Arial10, Element.ALIGN_CENTER))

        'TablaPie.AddCell(Crear_Celda("RECEIVED BY", Arial12, Element.ALIGN_LEFT))
        'TablaPie.AddCell(Crear_Celda("DATE", Arial12, Element.ALIGN_LEFT))

        Obtener_Pie_CPX = TablaPie
    End Function

    '***********************************************************
    ' CREAR PDF DE PIDOBOX
    '***********************************************************
    Function Crear_Reporte_Pdf_PidoBox(ByVal Guia As String, ByRef NombreArchivo As String, ByRef Mensaje As String, ByVal CodigoImportador As Integer) As Boolean

        Try
            Dim TablaGeneral As PdfPTable = New PdfPTable(1)
            TablaGeneral.DefaultCell.Border = Rectangle.NO_BORDER
            TablaGeneral.WidthPercentage = 100

            Celda = New PdfPCell(Obtener_Encabezado_PidoBox(Guia, CodigoImportador))
            Celda.HorizontalAlignment = Element.ALIGN_LEFT
            Celda.Border = Rectangle.NO_BORDER
            Celda.BorderWidthBottom = 1
            Celda.PaddingBottom = 2
            Celda.PaddingTop = 3
            TablaGeneral.AddCell(Celda)

            'Linea
            Frase = New Phrase()
            SubParrafo1 = New Chunk(vbNewLine, Arial10)
            Frase.Add(SubParrafo1)
            TablaGeneral.AddCell(Frase)

            Celda = New PdfPCell(Obtener_Titulo_Destino_PidoBox)
            Celda.HorizontalAlignment = Element.ALIGN_CENTER
            Celda.VerticalAlignment = Element.ALIGN_BOTTOM
            Celda.Border = Rectangle.NO_BORDER
            TablaGeneral.AddCell(Celda)


            Celda = New PdfPCell(Obtener_Direccion_Empresa_PidoBox(EnviadoPor))
            Celda.HorizontalAlignment = Element.ALIGN_CENTER
            Celda.VerticalAlignment = Element.ALIGN_CENTER
            Celda.Border = Rectangle.NO_BORDER
            TablaGeneral.AddCell(Celda)

            'Linea
            Frase = New Phrase()
            SubParrafo1 = New Chunk(vbNewLine, Arial10)
            Frase.Add(SubParrafo1)
            TablaGeneral.AddCell(Frase)

            Celda = New PdfPCell(Obtener_Descripcion_PidoBox(Descripcion, Peso, Tracking))
            Celda.HorizontalAlignment = Element.ALIGN_LEFT
            Celda.Border = Rectangle.NO_BORDER
            TablaGeneral.AddCell(Celda)


            'documento tamaño carta
            Dim doc1 = New Document(iTextSharp.text.PageSize.LETTER)
            Dim writer As PdfWriter = PdfWriter.GetInstance(doc1, New FileStream("C:\inetpub\wwwroot\Sistema\CARGA\" + Tracking + ".pdf", FileMode.Create))
            NombreArchivo = "C:\inetpub\wwwroot\Sistema\CARGA\" + Tracking + ".pdf"
            doc1.Open()

            doc1.Add(TablaGeneral)

            doc1.Close()
            writer.Close()

            Return True
        Catch ex As Exception
            Mensaje = ex.Message.ToString
            Return False
        End Try
    End Function

    Function Obtener_Encabezado_PidoBox(ByVal Guia As String, ByVal CodigoImportador As Integer) As PdfPTable
        Dim TablaEncabezado As PdfPTable = New PdfPTable(3)
        Dim AnchoEncabezado() As Single = {35, 30, 35}

        TablaEncabezado.WidthPercentage = 100
        TablaEncabezado.SetWidths(AnchoEncabezado)
        TablaEncabezado.DefaultCell.Border = Rectangle.NO_BORDER

        'Celda 1
        'Dim fotoencabezado As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(imagepath + "/PidoBox.gif")
        Dim fotoencabezado As iTextSharp.text.Image

        If CodigoImportador = 6 Then 'Pidobox
            fotoencabezado = iTextSharp.text.Image.GetInstance(imagepath + "/PidoBox.gif")
        ElseIf CodigoImportador = 7 Then 'RapiPuerta
            fotoencabezado = iTextSharp.text.Image.GetInstance(imagepath + "/RapiPuerta.png")
        ElseIf CodigoImportador = 8 Then 'QS
            fotoencabezado = iTextSharp.text.Image.GetInstance(imagepath + "/qs.png")
        ElseIf CodigoImportador = 9 Then 'ICC
            fotoencabezado = iTextSharp.text.Image.GetInstance(imagepath + "/icc.png")
        Else 'YoCargo
            fotoencabezado = iTextSharp.text.Image.GetInstance(imagepath + "/YoCargo.png")
        End If

        fotoencabezado.ScalePercent(48.0F)
        fotoencabezado.SpacingAfter = 5 '10

        Celda = New PdfPCell(fotoencabezado)
        Celda.HorizontalAlignment = Element.ALIGN_LEFT
        Celda.VerticalAlignment = Element.ALIGN_TOP
        Celda.Border = Rectangle.NO_BORDER
        TablaEncabezado.AddCell(Celda)

        'Celda 2
        Frase = New Phrase()
        SubParrafo1 = New Chunk(Linea1AWB + vbNewLine, Arial10)
        SubParrafo2 = New Chunk(Linea2AWB + vbNewLine, Arial10)
        SubParrafo3 = New Chunk(Linea3AWB + vbNewLine, Arial10)
        SubParrafo4 = New Chunk("" + vbNewLine, Arial10)

        Frase.Add(SubParrafo1)
        Frase.Add(SubParrafo2)
        Frase.Add(SubParrafo3)
        Frase.Add(SubParrafo4)
        Frase.Add(SubParrafo4)

        Celda = New PdfPCell(Frase)
        Celda.HorizontalAlignment = Element.ALIGN_LEFT
        Celda.VerticalAlignment = Element.ALIGN_BOTTOM
        Celda.Border = Rectangle.NO_BORDER
        TablaEncabezado.AddCell(Celda)

        'Celda 3
        Dim TablaCelda3 As PdfPTable = New PdfPTable(1)
        Dim AnchoCelda3() As Single = {35}

        TablaCelda3.WidthPercentage = 100
        TablaCelda3.SetWidths(AnchoCelda3)
        TablaCelda3.DefaultCell.Border = Rectangle.NO_BORDER

        'Celda 3.1
        Frase = New Phrase()
        If CodigoImportador = 6 Then 'Pidobox
            SubParrafo1 = New Chunk("GUÍA POBOX" + vbNewLine, Arial10)
        ElseIf CodigoImportador = 7 Then 'RapiPuerta
            SubParrafo1 = New Chunk("GUÍA" + vbNewLine, Arial10)
        ElseIf CodigoImportador = 8 Then 'QS
            SubParrafo1 = New Chunk("GUÍA QS" + vbNewLine, Arial10)
        ElseIf CodigoImportador = 9 Then 'ICC
            SubParrafo1 = New Chunk("GUÍA ICC" + vbNewLine, Arial10)
        Else 'YoCargo
            SubParrafo1 = New Chunk("GUÍA" + vbNewLine, Arial10)
        End If

        Frase.Add(SubParrafo1)

        Celda = New PdfPCell(Frase)
        Celda.HorizontalAlignment = Element.ALIGN_CENTER
        Celda.Border = Rectangle.NO_BORDER
        TablaCelda3.AddCell(Celda)

        'Celda 3.2
        'Dim bm As System.Drawing.Bitmap = CodigoBarras.CodeCodABAR("A0" + Guia + "B", False, 12) 'Sólo números
        Dim bm As System.Drawing.Bitmap = CodigoBarras.Code128(Guia, 9, False, 30, , False) 'Números y letras

        Dim ImagenBarras As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(bm, System.Drawing.Imaging.ImageFormat.Png)

        'ImagenBarras.ScalePercent(150, 200) 'escala al tamaño de la imagen 

        Celda = New PdfPCell(ImagenBarras)
        Celda.HorizontalAlignment = Element.ALIGN_CENTER
        Celda.VerticalAlignment = Element.ALIGN_CENTER
        Celda.Border = Rectangle.NO_BORDER
        Celda.PaddingRight = 10
        Celda.PaddingBottom = 1
        TablaCelda3.AddCell(Celda)

        'Celda 3.3
        Frase = New Phrase()
        SubParrafo1 = New Chunk(Guia + vbNewLine, Arial10)
        Frase.Add(SubParrafo1)

        Celda = New PdfPCell(Frase)
        Celda.HorizontalAlignment = Element.ALIGN_CENTER
        Celda.VerticalAlignment = Element.ALIGN_TOP
        Celda.Border = Rectangle.NO_BORDER
        TablaCelda3.AddCell(Celda)

        'Agrega a la celda 3
        Celda = New PdfPCell(TablaCelda3)
        Celda.HorizontalAlignment = Element.ALIGN_LEFT
        Celda.VerticalAlignment = Element.ALIGN_BOTTOM
        Celda.Border = Rectangle.NO_BORDER
        TablaEncabezado.AddCell(Celda)

        Obtener_Encabezado_PidoBox = TablaEncabezado
    End Function

    Function Obtener_Titulo_Destino_PidoBox() As PdfPTable
        Dim TablaTitulos As PdfPTable = New PdfPTable(2)
        Dim AnchoTitulos() As Single = {50, 50}
        TablaTitulos.WidthPercentage = 100
        TablaTitulos.SetWidths(AnchoTitulos)
        TablaTitulos.DefaultCell.Border = Rectangle.NO_BORDER

        TablaTitulos.AddCell(Crear_Celda("Envíado por (Remitente):", Arial10Negrita, Element.ALIGN_LEFT))
        TablaTitulos.AddCell(Crear_Celda("Para (Destinatario):", Arial10Negrita, Element.ALIGN_LEFT))

        Obtener_Titulo_Destino_PidoBox = TablaTitulos
    End Function

    Function Obtener_Direccion_Empresa_PidoBox(ByVal EnviadoPor As String) As PdfPTable
        'Tabla para el borde
        Dim TablaRedondeada1 As PdfPTable = New PdfPTable(1)
        TablaRedondeada1.DefaultCell.Border = PdfPCell.NO_BORDER
        TablaRedondeada1.DefaultCell.CellEvent = New RoundedBorder()
        TablaRedondeada1.DefaultCell.Padding = 5 '10

        Dim TablaDireccion As PdfPTable = New PdfPTable(2)
        Dim AnchoDireccion() As Single = {50, 50}
        TablaDireccion.WidthPercentage = 100
        TablaDireccion.SetWidths(AnchoDireccion)
        TablaDireccion.DefaultCell.Border = Rectangle.NO_BORDER


        TablaDireccion.AddCell(Crear_Celda("Nombre remitente o compañía: " + EnviadoPor, Arial9, Element.ALIGN_LEFT))
        TablaDireccion.AddCell(Crear_Celda("Nombre: Guatemala Digital S.A.", Arial9, Element.ALIGN_LEFT))

        TablaDireccion.AddCell(Crear_Celda("Dirección: USA", Arial9, Element.ALIGN_LEFT))
        TablaDireccion.AddCell(Crear_Celda("Dirección de Entrega: Ciudad, Guatemala City Guatemala", Arial9, Element.ALIGN_LEFT))

        TablaDireccion.AddCell(Crear_Celda(" ", Arial9, Element.ALIGN_LEFT))
        TablaDireccion.AddCell(Crear_Celda("Nit: 8316715-3", Arial9, Element.ALIGN_LEFT))

        TablaDireccion.AddCell(Crear_Celda("Teléfono: ", Arial9, Element.ALIGN_LEFT))
        TablaDireccion.AddCell(Crear_Celda("País: Guatemala", Arial9, Element.ALIGN_LEFT))

        TablaRedondeada1.AddCell(TablaDireccion)

        Obtener_Direccion_Empresa_PidoBox = TablaRedondeada1

    End Function

    Function Obtener_Descripcion_PidoBox(ByVal Descripcion As String, ByVal Peso As Decimal, ByVal Tracking As String) As PdfPTable

        Dim Fecha As Date
        Dim CadPeso, CadFreight As String
        Dim TablaDescripcion As PdfPTable = New PdfPTable(2)
        Dim AnchoDescripcion() As Single = {5, 95}
        TablaDescripcion.WidthPercentage = 100
        TablaDescripcion.SetWidths(AnchoDescripcion)
        TablaDescripcion.DefaultCell.Border = Rectangle.NO_BORDER

        Dim TrackingReducido, MensajeError, CadPiezas As String

        If Peso > 0 Then
            CadPeso = Format(Peso, "##,##0") & " Pounds"
            CadFreight = "$" + Format(Peso * 1.6, "##,##0.00")
        Else
            CadPeso = "NULL"
            CadFreight = "NULL"
        End If


        TrackingReducido = "" : MensajeError = "" : CadPiezas = "null"
        If Validar_Tracking(Tracking, TrackingReducido, MensajeError) = True Then

            If TrackingReducido <> "" Then
                Consulta = "select SUM(cantidad) from pedido where codigoderastreo like '%" & TrackingReducido & "%'"
                CadPiezas = CStr(mostrar2.retornarentero(Consulta, MyconString))

                Consulta = "select top(1) FechaDeEnvio from pedido where codigoderastreo like '%" & TrackingReducido & "%' order by FechaDeEnvio desc"
                Fecha = Obtener_Fecha_envio(Consulta)

            Else
                Consulta = "select SUM(cantidad) from pedido where codigoderastreo like '%" & Tracking & "%'"
                CadPiezas = CStr(mostrar2.retornarentero(Consulta, MyconString))

                Consulta = "select top(1) FechaDeEnvio from pedido where codigoderastreo like '%" & Tracking & "%' order by FechaDeEnvio desc"
                Fecha = Obtener_Fecha_envio(Consulta)

            End If

            If CadPiezas = "0" Then
                CadPiezas = "null"
            End If

            If DateAdd(DateInterval.Day, 20, Fecha) > Date.Now Then
                Fecha = Date.Now
            Else
                Fecha = DateAdd(DateInterval.Day, 20, Fecha)
            End If

        End If

        TablaDescripcion.AddCell(Crear_Celda_Padding7(" ", Arial5, Element.ALIGN_LEFT))
        TablaDescripcion.AddCell(Crear_Celda_Padding7(" ", Arial5, Element.ALIGN_LEFT))

        TablaDescripcion.AddCell(" ")
        TablaDescripcion.AddCell(Crear_Celda_Padding7("Fecha: " + Obtener_Fecha(Fecha), Arial10, Element.ALIGN_LEFT))

        TablaDescripcion.AddCell(" ")
        TablaDescripcion.AddCell(Crear_Celda_Padding7("Origen: ", Arial10, Element.ALIGN_LEFT))

        TablaDescripcion.AddCell(" ")
        TablaDescripcion.AddCell(Crear_Celda_Padding7("Vuelo Número: ", Arial10, Element.ALIGN_LEFT))

        TablaDescripcion.AddCell(" ")
        TablaDescripcion.AddCell(Crear_Celda_Padding7("Destino: " + "Guatemala", Arial10, Element.ALIGN_LEFT))

        TablaDescripcion.AddCell(" ")
        TablaDescripcion.AddCell(Crear_Celda_Padding7("Descrición de contenido: " + Descripcion, Arial10, Element.ALIGN_LEFT))

        TablaDescripcion.AddCell(" ")
        TablaDescripcion.AddCell(Crear_Celda_Padding7("Piezas: " + CadPiezas, Arial10, Element.ALIGN_LEFT))

        TablaDescripcion.AddCell(" ")
        TablaDescripcion.AddCell(Crear_Celda_Padding7("Peso: " + CadPeso, Arial10, Element.ALIGN_LEFT))

        TablaDescripcion.AddCell(" ")
        TablaDescripcion.AddCell(Crear_Celda_Padding7("Tracking: " + Tracking, Arial10, Element.ALIGN_LEFT))

        TablaDescripcion.AddCell(" ")
        TablaDescripcion.AddCell(Crear_Celda_Padding7("Carga: " + CadFreight, Arial10, Element.ALIGN_LEFT))

        Obtener_Descripcion_PidoBox = TablaDescripcion

    End Function

    'Function Obtener_Pie_PidoBox(ByVal Tracking As String) As PdfPTable
    '    Dim Fecha As Date
    '    Dim TrackingReducido, MensajeError As String

    '    Dim TablaPie As PdfPTable = New PdfPTable(3)
    '    Dim AnchoPie() As Single = {30, 42, 28}
    '    TablaPie.WidthPercentage = 100
    '    TablaPie.SetWidths(AnchoPie)
    '    TablaPie.DefaultCell.Border = Rectangle.NO_BORDER

    '    TrackingReducido = "" : MensajeError = ""
    '    If Validar_Tracking(Tracking, TrackingReducido, MensajeError) = True Then

    '        If TrackingReducido <> "" Then
    '            Consulta = "select top(1) FechaDeEnvio from pedido where codigoderastreo like '%" & TrackingReducido & "%' order by FechaDeEnvio desc"
    '            Fecha = Obtener_Fecha_envio(Consulta)
    '        Else
    '            Consulta = "select top(1) FechaDeEnvio from pedido where codigoderastreo like '%" & Tracking & "%' order by FechaDeEnvio desc"
    '            Fecha = Obtener_Fecha_envio(Consulta)
    '        End If

    '    End If

    '    If DateAdd(DateInterval.Day, 20, Fecha) > Date.Now Then
    '        Fecha = Date.Now
    '    Else
    '        Fecha = DateAdd(DateInterval.Day, 20, Fecha)
    '    End If

    '    Frase = New Phrase()
    '    SubParrafo1 = New Chunk("" + vbNewLine, Arial10Negrita)
    '    SubParrafo2 = New Chunk("___________________________" + vbNewLine, Arial10)
    '    SubParrafo3 = New Chunk("RECEIVED BY", Arial10)
    '    Frase.Add(SubParrafo1)
    '    Frase.Add(SubParrafo2)
    '    Frase.Add(SubParrafo3)
    '    Celda = New PdfPCell(Frase)
    '    Celda.HorizontalAlignment = Element.ALIGN_CENTER
    '    Celda.Border = Rectangle.NO_BORDER
    '    TablaPie.AddCell(Celda)

    '    Frase = New Phrase()
    '    SubParrafo1 = New Chunk("MIAMI, FLORIDA " + Obtener_Fecha(Fecha) + "  " + vbNewLine, Arial10Negrita)
    '    SubParrafo2 = New Chunk("______________________________" + vbNewLine, Arial10)
    '    SubParrafo3 = New Chunk("DATE", Arial10)
    '    Frase.Add(SubParrafo1)
    '    Frase.Add(SubParrafo2)
    '    Frase.Add(SubParrafo3)
    '    Celda = New PdfPCell(Frase)
    '    Celda.HorizontalAlignment = Element.ALIGN_CENTER
    '    Celda.Border = Rectangle.NO_BORDER
    '    TablaPie.AddCell(Celda)

    '    TablaPie.AddCell("")

    '    'TablaPie.AddCell("")
    '    'TablaPie.AddCell(Crear_Celda("MIAMI, FLORIDA 09/14/2015", Arial10Negrita, Element.ALIGN_CENTER))


    '    'TablaPie.AddCell(Crear_Celda("______________________________" + vbNewLine + "RECEIVED BY", Arial10, Element.ALIGN_CENTER))
    '    'TablaPie.AddCell(Crear_Celda("______________________________" + vbNewLine + "DATE", Arial10, Element.ALIGN_CENTER))

    '    'TablaPie.AddCell(Crear_Celda("RECEIVED BY", Arial12, Element.ALIGN_LEFT))
    '    'TablaPie.AddCell(Crear_Celda("DATE", Arial12, Element.ALIGN_LEFT))

    '    Obtener_Pie_PidoBox = TablaPie
    'End Function


    '***********************************************************
    ' FUNCIONES GENERALES
    '***********************************************************
    '>> Gloria Tarea 234 21-feb-2018
    Sub Obtener_Datos_Importador(ByVal Guia As String, ByRef EnviarExencion As Boolean)

        Consulta = "select top 1 case when ISNULL(pe.CodigoImportador,0) > 0 then pe.CodigoImportador else isnull(pa.CodigoImportador,0) end as CodigoImportador, i.Nombre, i.Linea1AWB, i.Linea2AWB, i.Linea3AWB, " &
            "case when CHARINDEX('NO-Enviar-Exención', Observaciones) + CHARINDEX('NO-Enviar-Exencion', Observaciones) > 0 then Convert(bit,0) else Convert(bit,1) end as EnviarExencion " &
            "from Pedido pe, Paquete pa, Importador i where pe.CodigoPaquete = pa.CodigoPaquete And GuiaAerea = '" & Guia & "' " &
            "And pe.CodigoImportador = i.CodigoImportador And Cantidad > 0 And CodigoEstadoPedido <> 4  order by 1 desc"

        Using mySqlConnection2 As New System.Data.SqlClient.SqlConnection(MyconString)
            mySqlConnection2.Open()
            Dim mySqlCommand2 As New System.Data.SqlClient.SqlCommand(Consulta, mySqlConnection2)
            Dim myDataReader2 As Data.SqlClient.SqlDataReader
            myDataReader2 = mySqlCommand2.ExecuteReader()


            Do While myDataReader2.Read()
                If myDataReader2.IsDBNull(0) = False Then
                    CodigoImportador = myDataReader2.GetInt32(0)
                End If
                NombreImportador = ""
                If myDataReader2.IsDBNull(1) = False Then
                    NombreImportador = myDataReader2.GetString(1)
                End If
                Linea1AWB = ""
                If myDataReader2.IsDBNull(2) = False Then
                    Linea1AWB = myDataReader2.GetString(2)
                End If
                Linea2AWB = ""
                If myDataReader2.IsDBNull(3) = False Then
                    Linea2AWB = myDataReader2.GetString(3)
                End If
                Linea3AWB = ""
                If myDataReader2.IsDBNull(4) = False Then
                    Linea3AWB = myDataReader2.GetString(4)
                End If
                EnviarExencion = False
                If myDataReader2.IsDBNull(5) = False Then
                    EnviarExencion = myDataReader2.GetBoolean(5)
                End If
            Loop
            myDataReader2.Close()
            mySqlConnection2.Close()
        End Using

    End Sub

    Function Validar_Tracking(ByVal Tracking As String, ByRef TrackingReducido As String, ByRef MensajeError As String) As Boolean
        Dim Exito As Boolean

        Exito = True
        MensajeError = ""
        TrackingReducido = ""

        If Len(Tracking) > 60 Then
            MensajeError = "Tracking demasiado largo"
            Exito = False
        Else
            'verifica si existe el tracking ingresado
            Consulta = "select Count(*)  from Pedido where (CodigoDeRastreo = '" & Tracking & "' or CodigoDeRastreo like '%" & Tracking & "%') and CodigoEstadoPedido <> 4 "
            If mostrar2.retornarentero(Consulta, MyconString) = 0 Then 'no hay tracking
                'verifica si el tracking es númerico (puede ser de Fedex)
                If IsNumeric(Tracking) = False Then
                    MensajeError = "El tracking no existe"
                    Exito = False
                Else 'toma los últimos 12 dígitos del tracking (por si es de Fedex)
                    TrackingReducido = Microsoft.VisualBasic.Right(Tracking, 12)
                    'verifica si existe el tracking 
                    Consulta = "select Count(*)  from Pedido where (CodigoDeRastreo = '" & TrackingReducido & "' or CodigoDeRastreo like '%" & TrackingReducido & "%')  and CodigoEstadoPedido <> 4 "
                    If mostrar2.retornarentero(Consulta, MyconString) = 0 Then 'no hay tracking
                        MensajeError = "El tracking no existe"
                        Exito = False
                    Else
                        'verifica si se encontró más de 1 tracking
                        Consulta = "select COUNT(*) from (select distinct(CodigoDeRastreo) from Pedido where (CodigoDeRastreo = '" & TrackingReducido & "' or CodigoDeRastreo like '%" & TrackingReducido & "%') and CodigoEstadoPedido <> 4 ) t"
                        If mostrar2.retornarentero(Consulta, MyconString) > 1 Then
                            Consulta = "select COUNT(*) from (select distinct(CodigoPaquete) from Pedido where (CodigoDeRastreo = '" & Tracking & "' or CodigoDeRastreo like '%" & Tracking & "%') and CodigoEstadoPedido <> 4 ) t"
                            If mostrar2.retornarentero(Consulta, MyconString) > 1 Then
                                MensajeError = "Se encontró más de 1 tracking"
                                Exito = False
                            End If
                        End If

                    End If
                End If

            Else
                'verifica si se encontró más de 1 tracking
                Consulta = "select COUNT(*) from (select distinct(CodigoDeRastreo) from Pedido where (CodigoDeRastreo = '" & Tracking & "' or CodigoDeRastreo like '%" & Tracking & "%') and CodigoEstadoPedido <> 4 ) t"
                If mostrar2.retornarentero(Consulta, MyconString) > 1 Then
                    Consulta = "select COUNT(*) from (select distinct(CodigoPaquete) from Pedido where (CodigoDeRastreo = '" & Tracking & "' or CodigoDeRastreo like '%" & Tracking & "%') and CodigoEstadoPedido <> 4 ) t"
                    If mostrar2.retornarentero(Consulta, MyconString) > 1 Then
                        MensajeError = "Se encontró más de 1 tracking"
                        Exito = False
                    End If
                End If
            End If 'no hay tracking

        End If 'longitud > 40

        Validar_Tracking = Exito
    End Function

    Sub Obtener_Datos_Guia(
        ByVal Guia As String, 
        ByRef CodigoPaquete As Integer, 
        ByRef EnviadoPor As String, 
        ByRef Descripcion As String, 
        ByRef Peso As Decimal, 
        ByRef Tracking As String
        )

        'obtener producto
        Consulta = "select CodigoPaquete, EnviadoPor, Descripcion, Peso, CodigoDeRastreo from Paquete where GuiaAerea = '" & Guia & "'"

        Using mySqlConnection2 As New System.Data.SqlClient.SqlConnection(MyconString)
            mySqlConnection2.Open()
            Dim mySqlCommand2 As New System.Data.SqlClient.SqlCommand(Consulta, mySqlConnection2)
            Dim myDataReader2 As Data.SqlClient.SqlDataReader
            myDataReader2 = mySqlCommand2.ExecuteReader()


            Do While myDataReader2.Read()
                If myDataReader2.IsDBNull(0) = False Then
                    CodigoPaquete = myDataReader2.GetInt32(0)
                End If
                EnviadoPor = ""
                If myDataReader2.IsDBNull(1) = False Then
                    EnviadoPor = myDataReader2.GetString(1)
                End If
                Descripcion = ""
                If myDataReader2.IsDBNull(2) = False Then
                    Descripcion = myDataReader2.GetString(2)
                End If
                If myDataReader2.IsDBNull(3) = False Then
                    Peso = myDataReader2.GetDecimal(3)
                End If
                Tracking = ""
                If myDataReader2.IsDBNull(4) = False Then
                    Tracking = myDataReader2.GetString(4)
                End If

            Loop

            myDataReader2.Close()
            mySqlConnection2.Close()
        End Using

    End Sub

    Function Obtener_Fecha(ByVal Fecha As Date) As String
        Dim Cadena As String
        Dim Dia, Mes, Anio As Integer
        Dim NombreMes, CadDia, CadMes As String

        Cadena = ""
        NombreMes = ""
        Dia = Fecha.Day
        Mes = Fecha.Month
        Anio = Fecha.Year


        If Dia < 10 Then
            CadDia = "0" + CStr(Dia)
        Else
            CadDia = CStr(Dia)
        End If

        If Mes < 10 Then
            CadMes = "0" + CStr(Mes)
        Else
            CadMes = CStr(Mes)
        End If

        Cadena = CadDia & "/" & CadMes & "/" & CStr(Anio)
        Obtener_Fecha = Cadena

    End Function

    Function Obtener_Fecha_envio(ByVal Consulta As String) As Date
        Dim Fecha As Date

        Using mySqlConnection2 As New System.Data.SqlClient.SqlConnection(MyconString)
            mySqlConnection2.Open()
            Dim mySqlCommand2 As New System.Data.SqlClient.SqlCommand(Consulta, mySqlConnection2)
            Dim myDataReader2 As Data.SqlClient.SqlDataReader
            myDataReader2 = mySqlCommand2.ExecuteReader()


            Do While myDataReader2.Read()
                If myDataReader2.IsDBNull(0) = False Then
                    Fecha = myDataReader2.GetDateTime(0)
                Else
                    Fecha = Date.Now
                End If

            Loop

            myDataReader2.Close()
            mySqlConnection2.Close()
        End Using

        Obtener_Fecha_envio = Fecha

    End Function

End Class

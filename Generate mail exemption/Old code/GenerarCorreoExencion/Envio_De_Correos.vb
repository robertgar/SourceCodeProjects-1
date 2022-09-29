Imports System.Security.Cryptography
Imports System.Text
Imports System
Imports System.IO
'Imports Amazon.SimpleNotificationService
'Imports Amazon.SimpleNotificationService.Model


Public Class Envio_De_Correos

    Function Enviar_Correo_Oferta_Notificacion(ByVal DireccionCorreoOrigen As String, ByVal DireccionCorreoDestino As String, ByVal Titulo As String, ByVal Contenido As String, ByVal CorreosAlternos As String) As Boolean
        Dim Mensaje As String
        Dim ImagenEncabezado As String
        Dim ImagenBarra As String
        Dim ImagenDesuscripcion As String
        Dim Texto() As String
        Dim correo As New System.Net.Mail.MailMessage
        Dim Exito As Boolean

        Exito = True
        Try

            correo.From = New System.Net.Mail.MailAddress(DireccionCorreoOrigen, "GuatemalaDigital.com")
            'correo.Bcc.Add("info@guatemaladigital.com")

            If CorreosAlternos <> "" Then
                Texto = Split(CorreosAlternos, ",")
                For i = 0 To Texto.Length - 1
                    correo.Bcc.Add(Texto(i))
                Next
            End If

            correo.To.Add(DireccionCorreoDestino) 'correo.To.Add(txtCorreo.Text)

            correo.Subject = Titulo

            ImagenEncabezado = "http://www.guatemaladigital.com//images/poster_correo.jpg"
            ImagenBarra = "http://www.guatemaladigital.com//images/CorreoEncabezado.jpg"
            ImagenDesuscripcion = "http://www.guatemaladigital.com//images/Desuscripcion.png"

            'Mensaje = "<table style=""width:800px; border-width:1px;  border-style:solid""> " & _
            '"<tr> " & _
            '"<td align=""left""> " & _
            '"<img src='" & ImagenEncabezado & "' alt=""cargando imagen"" width=""800px""> " & _
            '"<img src='" & ImagenBarra & "' alt=""cargando imagen"" width=""800px""> " & _
            ' Contenido & _
            '"<img src='" & ImagenBarra & "' alt=""cargando imagen"" width=""800px""> " & _
            '" </td> " & _
            '"</tr> " & _
            '"</table> "

            Mensaje = Contenido


            correo.Body = Mensaje

            correo.IsBodyHtml = True
            correo.Priority = System.Net.Mail.MailPriority.Normal
            Dim smtp As New System.Net.Mail.SmtpClient
            smtp.Host = "email-smtp.us-east-1.amazonaws.com"
            smtp.Port = 25
            smtp.EnableSsl = True
            smtp.Credentials = New System.Net.NetworkCredential("AKIAJFKC2GTUH7X3AISA", "Ao0QQUw2/KRPfQLSH5Msg+zmx0iXe+1hwThQXqoHxibh")
            smtp.Send(correo)

            Exito = True

        Catch ex As Exception
            Exito = False
        End Try

        Enviar_Correo_Oferta_Notificacion = Exito
    End Function


    Sub Enviar_Correo(ByVal DireccionCorreoOrigen As String, ByVal DireccionCorreoDestino As String, ByVal Titulo As String, ByVal Contenido As String, ByVal CorreosAlternos As String)
        Dim Mensaje As String
        Dim ImagenEncabezado As String
        Dim ImagenBarra As String
        Dim ImagenDesuscripcion As String
        Dim Texto() As String
        Dim correo As New System.Net.Mail.MailMessage


        correo.From = New System.Net.Mail.MailAddress(DireccionCorreoOrigen, "GuatemalaDigital.com")
        'correo.Bcc.Add("info@guatemaladigital.com")

        If CorreosAlternos <> "" Then
            Texto = Split(CorreosAlternos, ",")
            For i = 0 To Texto.Length - 1
                correo.Bcc.Add(Texto(i))
            Next
        End If

        correo.To.Add(DireccionCorreoDestino) 'correo.To.Add(txtCorreo.Text)

        correo.Subject = Titulo

        ImagenEncabezado = "http://www.guatemaladigital.com//images/poster_correo.jpg"
        ImagenBarra = "http://www.guatemaladigital.com//images/CorreoEncabezado.jpg"
        ImagenDesuscripcion = "http://www.guatemaladigital.com//images/Desuscripcion.png"


        'Mensaje = Contenido
        Mensaje = "<table style=""width:800px; border-width:1px;  border-style:solid""> " &
            "<tr> " &
            "<td align=""left""> " &
            "<img src='" & ImagenEncabezado & "' alt=""cargando imagen"" width=""800px""> " &
            "<img src='" & ImagenBarra & "' alt=""cargando imagen"" width=""800px""> " &
             Contenido &
            "<img src='" & ImagenBarra & "' alt=""cargando imagen"" width=""800px""> " &
            " </td> " &
            "</tr> " &
            "</table> "


        correo.Body = Mensaje

        correo.IsBodyHtml = True
        correo.Priority = System.Net.Mail.MailPriority.Normal
        Dim smtp As New System.Net.Mail.SmtpClient
        smtp.Host = "email-smtp.us-east-1.amazonaws.com"
        smtp.Port = 25
        smtp.EnableSsl = True
        smtp.Credentials = New System.Net.NetworkCredential("AKIAJFKC2GTUH7X3AISA", "Ao0QQUw2/KRPfQLSH5Msg+zmx0iXe+1hwThQXqoHxibh")
        smtp.Send(correo)



    End Sub

    Sub Enviar_Correo_Con_Attachment(ByVal DireccionCorreoOrigen As String, ByVal DireccionCorreoDestino As String, ByVal Titulo As String, ByVal Contenido As String, ByVal CorreosAlternos As String, ByVal ArchivosIncrustados As String)
        Dim Mensaje As String
        Dim ImagenEncabezado As String
        Dim ImagenBarra As String
        Dim ImagenDesuscripcion As String
        Dim Texto(), Archivos() As String
        Dim correo As New System.Net.Mail.MailMessage


        correo.From = New System.Net.Mail.MailAddress(DireccionCorreoOrigen, "GuatemalaDigital.com")
        'correo.Bcc.Add("info@guatemaladigital.com")

        If CorreosAlternos <> "" Then
            Texto = Split(CorreosAlternos, ",")
            For i = 0 To Texto.Length - 1
                correo.Bcc.Add(Texto(i))
            Next
        End If

        correo.To.Add(DireccionCorreoDestino) 'correo.To.Add(txtCorreo.Text)

        correo.Subject = Titulo

        If ArchivosIncrustados <> "" Then
            Archivos = Split(ArchivosIncrustados, ",")
            For i = 0 To Archivos.Length - 1
                'correo.Attachments.Add(New Attachment("c:\\temp\\example.txt"))
                correo.Attachments.Add(New System.Net.Mail.Attachment(Archivos(i)))
            Next
        End If

        ImagenEncabezado = "http://www.guatemaladigital.com//images/poster_correo.jpg"
        ImagenBarra = "http://www.guatemaladigital.com//images/CorreoEncabezado.jpg"
        ImagenDesuscripcion = "http://www.guatemaladigital.com//images/Desuscripcion.png"

        'Mensaje = "<table style=""width:800px; border-width:1px;  border-style:solid""> " & _
        '"<tr> " & _
        '"<td align=""left""> " & _
        '"<img src='" & ImagenEncabezado & "' alt=""cargando imagen"" width=""800px""> " & _
        '"<img src='" & ImagenBarra & "' alt=""cargando imagen"" width=""800px""> " & _
        ' Contenido & _
        '"<img src='" & ImagenBarra & "' alt=""cargando imagen"" width=""800px""> " & _
        '" </td> " & _
        '"</tr> " & _
        '"</table> "

        Mensaje = Contenido


        correo.Body = Mensaje

        correo.IsBodyHtml = True
        correo.Priority = System.Net.Mail.MailPriority.Normal
        Dim smtp As New System.Net.Mail.SmtpClient
        smtp.Host = "email-smtp.us-east-1.amazonaws.com"
        smtp.Port = 25
        smtp.EnableSsl = True
        smtp.Credentials = New System.Net.NetworkCredential("AKIAJFKC2GTUH7X3AISA", "Ao0QQUw2/KRPfQLSH5Msg+zmx0iXe+1hwThQXqoHxibh")
        smtp.Send(correo)


    End Sub




    Function Enviar_Correo_Con_Attachment_Verificar_Error(ByVal DireccionCorreoOrigen As String, ByVal DireccionCorreoDestino As String, ByVal Titulo As String, ByVal Contenido As String, ByVal CorreosAlternos As String, ByVal ArchivosIncrustados As String) As Boolean
        Dim Mensaje As String
        Dim ImagenEncabezado As String
        Dim ImagenBarra As String
        Dim ImagenDesuscripcion As String
        Dim Texto(), Archivos() As String
        Dim correo As New System.Net.Mail.MailMessage

        Try


            correo.From = New System.Net.Mail.MailAddress(DireccionCorreoOrigen, "GuatemalaDigital.com")
            'correo.Bcc.Add("info@guatemaladigital.com")

            If CorreosAlternos <> "" Then
                Texto = Split(CorreosAlternos, ",")
                For i = 0 To Texto.Length - 1
                    correo.Bcc.Add(Texto(i))
                Next
            End If

            correo.To.Add(DireccionCorreoDestino) 'correo.To.Add(txtCorreo.Text)

            correo.Subject = Titulo

            If ArchivosIncrustados <> "" Then
                Archivos = Split(ArchivosIncrustados, ",")
                For i = 0 To Archivos.Length - 1
                    'correo.Attachments.Add(New Attachment("c:\\temp\\example.txt"))
                    correo.Attachments.Add(New System.Net.Mail.Attachment(Archivos(i)))
                Next
            End If

            ImagenEncabezado = "http://www.guatemaladigital.com//images/poster_correo.jpg"
            ImagenBarra = "http://www.guatemaladigital.com//images/CorreoEncabezado.jpg"
            ImagenDesuscripcion = "http://www.guatemaladigital.com//images/Desuscripcion.png"

            'Mensaje = "<table style=""width:800px; border-width:1px;  border-style:solid""> " & _
            '"<tr> " & _
            '"<td align=""left""> " & _
            '"<img src='" & ImagenEncabezado & "' alt=""cargando imagen"" width=""800px""> " & _
            '"<img src='" & ImagenBarra & "' alt=""cargando imagen"" width=""800px""> " & _
            ' Contenido & _
            '"<img src='" & ImagenBarra & "' alt=""cargando imagen"" width=""800px""> " & _
            '" </td> " & _
            '"</tr> " & _
            '"</table> "

            Mensaje = Contenido


            correo.Body = Mensaje

            correo.IsBodyHtml = True
            correo.Priority = System.Net.Mail.MailPriority.Normal
            Dim smtp As New System.Net.Mail.SmtpClient
            smtp.Host = "email-smtp.us-east-1.amazonaws.com"
            smtp.Port = 25
            smtp.EnableSsl = True
            smtp.Credentials = New System.Net.NetworkCredential("AKIAJFKC2GTUH7X3AISA", "Ao0QQUw2/KRPfQLSH5Msg+zmx0iXe+1hwThQXqoHxibh")
            smtp.Send(correo)

            Enviar_Correo_Con_Attachment_Verificar_Error = True

        Catch ex As Exception
            Enviar_Correo_Con_Attachment_Verificar_Error = False
        End Try


    End Function

    Sub Enviar_Correo_Rastreo(ByVal codigoventa As String, ByVal MyConString As String)
        Dim Contenido, Estilo, Foto, Consulta, PaginaProducto, CadenaUrl, NombreProducto As String
        Dim CodigoProducto, FechaConfirmacion, NombreCliente, CorreoCliente, Telefonos, Telefono1, Telefono2, DireccionDeEntrega As String
        Dim CodigoFormaPago, FormaPago, Monto, Envio, ServicioPago, Total, Factura, Nit, DireccionFactura, Titulo As String
        Dim Cuotas, EsPedido, IndiceNit, IndiceDir As Integer
        Dim total_pasos, codigo_estado_entrega, pasos_avanzados, indice As Integer
        Dim texto_status, texto_estado, div_visible, ancho, cadena, guia, fondo, divs_pasos, clase_spanrastreo, clase_marcarastreo, span_estado, span_estado2 As String
        Dim NombreRecibido, EmpresaDeEntrega, FechaEntrega, DiasAereo, NombreFactura, entregado, recibido As String
        Dim avance_marca As Double
        Dim cadenas() As String
        Dim mostrar As New Cargar
        Dim cargar As New Cargar
        Dim dtDatos, dt_estados_entrega, dtEntrega As New DataTable

        Consulta = "select Valor from Parametro where CodigoParametro = 8"
        DiasAereo = cargar.retornarcadena(Consulta, MyConString)

        Consulta = "SELECT v.codigoventa, v.CodigoProducto, Convert(varchar(10),CONVERT(date,v.FechaConfirmacion,106),103) as FechaConfirmacion, "
        Consulta += "convert(varchar,v.Cantidad)+' '+v.NombreProducto as NombreProducto, v.NombreCliente, v.CorreoCliente, v.Telefonos, ISNULL(cast(v.espedido as int),0) as EsPedido, "
        Consulta += "(v.DireccionDeEntrega+' Departamento: '+(select nombre from Departamento where codigodepartamento = v.codigodepartamento)+' Municipio: '+(select Nombre from Municipio where codigomunicipio = v.codigomunicipio)) as DireccionDeEntrega, "
        Consulta += "v.CodigoFormaDePago, (select Nombre from FormaDePago where CodigoFormaDePago = v.codigoformadepago) as FormaDePago, v.Monto, v.Envio, isnull(v.Cuotas,1) as cuotas, "
        Consulta += "case when v.CodigoFormaDePago = 1 then v.ServicioPagoEfectivo when v.CodigoFormaDePago = 3 then (v.MontoCuota-v.Monto) else 0.00 end as ServicioPago, "
        Consulta += "(v.Monto + v.Envio + (case when v.CodigoFormaDePago = 1 then v.ServicioPagoEfectivo when v.CodigoFormaDePago = 3 then (v.MontoCuota-v.Monto) else 0.00 end)) AS Total, "
        Consulta += "v.Factura, (select foto from Producto where codigoproducto = v.CodigoProducto) as foto, isnull(v.nombrerecibido,'') as NombreRecibido, "
        Consulta += "isnull(Convert(varchar(10),CONVERT(date,v.FechaEntrega,106),103),'') as FechaEntrega, "
        Consulta += "(select Nombre from EmpresaDeEntrega where CodigoEmpresaDeEntrega = v.CodigoEmpresaDeEntrega) as EmpresaDeEntrega  FROM venta v WHERE CodigoEstadoDeVenta = 1 and codigoventa = " + codigoventa
        cargar.ejecuta_query_dt(Consulta, dtDatos, MyConString)

        If dtDatos.Rows.Count <= 0 Then
            Exit Sub
        End If

        Cuotas = 0
        EsPedido = 0

        For Each fila As DataRow In dtDatos.Rows
            CodigoProducto = fila("CodigoProducto").ToString
            FechaConfirmacion = fila("FechaConfirmacion").ToString
            NombreProducto = fila("NombreProducto")
            NombreCliente = fila("NombreCliente").ToString
            CorreoCliente = fila("CorreoCliente").ToString
            Telefonos = fila("Telefonos").ToString
            EsPedido = CInt(fila("EsPedido").ToString)
            DireccionDeEntrega = fila("DireccionDeEntrega").ToString
            CodigoFormaPago = fila("CodigoFormaDePago").ToString
            FormaPago = fila("FormaDePago").ToString
            Monto = fila("Monto").ToString
            Envio = fila("Envio").ToString
            Cuotas = CInt(fila("cuotas").ToString)
            ServicioPago = fila("ServicioPago").ToString
            Total = fila("Total").ToString
            Factura = fila("Factura").ToString
            Foto = fila("Foto").ToString
            NombreRecibido = fila("NombreRecibido").ToString
            FechaEntrega = fila("FechaEntrega").ToString
            EmpresaDeEntrega = fila("EmpresaDeEntrega").ToString
        Next

        'Modificacion Hecha Por Diego
        If Telefonos.Contains("|") Then
            Telefono1 = Telefonos.Substring(0, Telefonos.IndexOf("|") - 1)
            Telefono2 = Telefonos.Substring(Telefonos.IndexOf("|") + 2)
        Else
            Telefono1 = Telefonos
            Telefono2 = ""
        End If
        'Fin Modificacion hecha por Diego

        IndiceNit = Factura.IndexOf("NIT:")
        IndiceDir = Factura.IndexOf("DIR:")

        If IndiceNit = -1 Then
            Nit = ""
            If IndiceDir = -1 Then
                DireccionFactura = ""
                NombreFactura = Factura.Substring(0)
            Else
                NombreFactura = Factura.Substring(0, IndiceDir)
                DireccionFactura = Factura.Substring(IndiceDir + 4)
            End If
        Else
            If IndiceDir = -1 Then
                NombreFactura = Factura.Substring(0, IndiceNit)
                Nit = Factura.Substring(IndiceNit + 4)
            Else
                NombreFactura = Factura.Substring(0, IndiceNit)
                Nit = Factura.Substring(IndiceNit + 4, IndiceDir - IndiceNit - 4)
                DireccionFactura = Factura.Substring(IndiceDir + 4)
            End If
        End If

        'Try
        '    Nit = Factura.Substring(InStr(1, Factura, "Nit:", CompareMethod.Text) + 4, InStr(1, Factura, "Dir:", CompareMethod.Text) - InStr(1, Factura, "Nit:", CompareMethod.Text) - 5)
        'Catch ex As Exception
        '    Nit = ""
        'End Try

        'Try
        '    DireccionFactura = Factura.Substring(Factura.LastIndexOf(":") + 1)
        'Catch ex As Exception
        '    DireccionFactura = ""
        'End Try

        Consulta = ""
        Consulta = "select '/' + " & mostrar.Reemplazar_Cadena_Url("c.Nombre") & " + '/' + " & mostrar.Reemplazar_Cadena_Url("p.Nombre") & " + '/' from Producto p, Categoria c " &
                     "where p.CodigoCategoria = c.CodigoCategoria " &
                     "and p.CodigoProducto = " & CodigoProducto

        CadenaUrl = mostrar.retornarcadena(Consulta, MyConString)
        CadenaUrl = mostrar.Longitud_Url(CadenaUrl)

        PaginaProducto = "http://www.guatemaladigital.com" & CadenaUrl & "Producto.aspx?Codigo=" & CodigoProducto

        If Trim(Foto) <> "" Then
            If InStr(Foto, "http") = 0 Then
                If Mid(Foto, 1, 1) <> "/" And Mid(Foto, 1, 1) <> "\" Then
                    Foto = "/" + Foto
                End If
                Foto = Replace(Foto, "\", "/")
                Foto = "http://www.guatemaladigital.com/" & Foto
            End If
        End If

        'Codigo que genera información de la barra de rastreo
        Consulta = ""
        Consulta = "select case when EsPedido IS Not null then 6 when EsPedido IS NULL and CodigoDeRastreo IS not null then 3 when EsPedido IS NULL and CodigoDeRastreo IS null then 2 end as pasos  from Venta where CodigoVenta = '" + codigoventa + "'"
        total_pasos = cargar.retornarentero(Consulta, MyConString)

        Consulta = ""
        Consulta = "select cast(codigoestadoentrega as int) from Venta where CodigoVenta = '" + codigoventa + "'"
        codigo_estado_entrega = cargar.retornarentero(Consulta, MyConString)

        Consulta = ""
        Consulta = "SELECT CodigoEstadoEntrega, nombre,  Etapa, ISNULL(CASE WHEN CodigoEstadoEntrega = 1 "
        Consulta += "THEN (SELECT CONVERT(varchar(10), CONVERT(date, fechadeorden, 106), 103) "
        Consulta += "FROM Pedido WHERE CodigoPedido = (SELECT CodigoPedido FROM VentaPedido "
        Consulta += "WHERE CodigoVenta = '" + codigoventa + "')) WHEN CodigoEstadoEntrega = 2 THEN (SELECT "
        Consulta += "CONVERT(varchar(10), CONVERT(date, FechaDeEnvio, 106), 103) FROM Pedido "
        Consulta += "WHERE CodigoPedido = (SELECT CodigoPedido FROM VentaPedido WHERE CodigoVenta = '" + codigoventa + "')) "
        Consulta += "WHEN CodigoEstadoEntrega = 3 THEN (SELECT CONVERT(varchar(10), CONVERT(date, Fecha, 106), 103) "
        Consulta += "FROM Paquete WHERE CodigoPaquete = (SELECT p.CodigoPaquete FROM Pedido p "
        Consulta += "INNER JOIN VentaPedido vp ON p.CodigoPedido = vp.CodigoPedido WHERE vp.CodigoVenta = '" + codigoventa + "')) "
        Consulta += "WHEN CodigoEstadoEntrega = 4 THEN (case when exists (SELECT FechaRecibida FROM Paquete WHERE CodigoPaquete = (SELECT p.CodigoPaquete "
        Consulta += "FROM Pedido p INNER JOIN VentaPedido vp ON p.CodigoPedido = vp.CodigoPedido WHERE vp.CodigoVenta = '" + codigoventa + "')) then "
        Consulta += "(SELECT CONVERT(varchar(10), CONVERT(date, FechaRecibida, 106), 103) FROM Paquete WHERE CodigoPaquete = (SELECT p.CodigoPaquete FROM Pedido p "
        Consulta += "INNER JOIN VentaPedido vp ON p.CodigoPedido = vp.CodigoPedido WHERE vp.CodigoVenta = '" + codigoventa + "')) "
        Consulta += "else (select CONVERT(varchar(10), CONVERT(date, fechaconfirmacion, 106), 103) from venta where codigoventa = '" + codigoventa + "') end) "
        Consulta += "WHEN CodigoEstadoEntrega = 5 THEN (SELECT CONVERT(varchar(10), CONVERT(date, Fecha, 106), 103) "
        Consulta += "FROM Factura WHERE CodigoFactura = (SELECT CodigoFactura FROM Venta WHERE CodigoVenta = '" + codigoventa + "')) "
        Consulta += "WHEN CodigoEstadoEntrega = 6 THEN (SELECT CONVERT(varchar(10), CONVERT(date, FechaEntrega, 106), 103) "
        Consulta += "FROM Venta WHERE CodigoVenta = '" + codigoventa + "') END, '') AS fecha FROM EstadoEntrega WHERE CodigoEstadoEntrega <= " + codigo_estado_entrega.ToString + ""

        cargar.ejecuta_query_dt(Consulta, dt_estados_entrega, MyConString)

        If dt_estados_entrega.Rows.Count <= 0 Then
            texto_status = "Sin Datos Para Rastreo"
            div_visible = "Visible = 'False'"
            Exit Sub
        End If

        cadena = ""
        For Each fila As DataRow In dt_estados_entrega.Rows
            guia = ""
            texto_status = fila("Etapa").ToString

            If fila("codigoestadoentrega").ToString = "1" Then
                cadena += fila("nombre").ToString.Substring(0, 12) + "<br/>" + fila("nombre").ToString.Substring(13) + "<br/>" + fila("fecha").ToString + ","
            ElseIf fila("codigoestadoentrega").ToString = "3" Then
                cadena += fila("nombre").ToString.Substring(0, 9) + "<br/>" + fila("nombre").ToString.Substring(10) + "<br/>" + fila("fecha").ToString + ","
            ElseIf fila("codigoestadoentrega").ToString = "5" Then
                Consulta = ""
                Consulta = "select codigoderastreo from Venta where CodigoVenta = " + codigoventa.ToString
                guia = cargar.retornarcadena(Consulta, MyConString)
                If (guia.Trim.Length > 0) And (guia.ToUpper <> "NULL") Then
                    guia = "<br/><a href='https://www.cargoexpreso.com/tracking/?guia=" + guia.ToString + "' target='_blank'>Rastrear " + guia.ToString + "</a>"
                Else
                    guia = ""
                End If
                cadena += fila("nombre").ToString.Substring(0, 11) + "<br/>" + fila("nombre").ToString.Substring(12) + "<br/>" + fila("fecha").ToString + guia.ToString + ","
            ElseIf fila("codigoestadoentrega").ToString = "6" Then
                Consulta = "select 'Recibe: <br/>'+v.NombreRecibido as recibe, 'Entrega: <br/>'+(select Nombre from EmpresaDeEntrega where CodigoEmpresaDeEntrega = v.CodigoEmpresaDeEntrega) as entrega from Venta v where v.CodigoVenta = " + codigoventa.ToString
                cargar.ejecuta_query_dt(Consulta, dtEntrega, MyConString)
                If dtEntrega.Rows.Count > 0 Then
                    recibido = "<br/>" + dtEntrega.Rows(0).Item(0).ToString
                    entregado = "<br/>" + dtEntrega.Rows(0).Item(1).ToString
                Else
                    recibido = ""
                    entregado = ""
                End If
                cadena += fila("nombre").ToString + "<br/>" + fila("fecha").ToString + entregado + recibido + ","
            Else
                cadena += fila("nombre").ToString + "<br/>" + fila("fecha").ToString + ","
            End If
        Next

        cadena = cadena.Substring(0, cadena.Length - 1)
        cadenas = Split(cadena, ",")

        Consulta = "select REPLACE(Nombre,'Bodega','En bodega') as Nombre from estadoentrega where CodigoEstadoEntrega = " + codigo_estado_entrega.ToString
        texto_estado = cargar.retornarcadena(Consulta, MyConString)

        Titulo = "Tu orden " + codigoventa + " de GuatemalaDigital esta " + texto_estado + "."

        'maneja la cantidad de divs que ha avanzado en la barra
        If total_pasos = 3 AndAlso codigo_estado_entrega >= 4 Then
            pasos_avanzados = codigo_estado_entrega - 4
        ElseIf total_pasos = 2 AndAlso codigo_estado_entrega = 4 Then
            pasos_avanzados = codigo_estado_entrega - 4
        ElseIf total_pasos = 2 AndAlso codigo_estado_entrega = 6 Then
            pasos_avanzados = codigo_estado_entrega - 5
        Else
            pasos_avanzados = codigo_estado_entrega - 1
        End If

        If codigo_estado_entrega = 6 Then
            clase_marcarastreo = "tracking-checkmark"
            clase_spanrastreo = "delivery-status-text status-green"
            fondo = "#218748"
            'texto_status = "Entregado"
        Else
            clase_marcarastreo = "tracking-circle"
            clase_spanrastreo = "delivery-status-text status-orange"
            fondo = "#f98c31"
            'texto_status = "En tránsito"
        End If

        If total_pasos = 6 Then
            ancho = "width: 20%;"
            span_estado = "style='margin-left: 11%; margin-right: -11%; float: left; position: relative; width: 14%; text-align: left; color:" + fondo + "; font-size: 14;'"
        ElseIf total_pasos = 3 Then
            ancho = "width: 50%;"
            span_estado = "style='margin-left: 12%; margin-right: 10%; float: left; position: relative; width: 11%; text-align: center; color:" + fondo + "; font-size: 14;'"
        ElseIf total_pasos = 2 Then
            ancho = "width: 100%;"
            span_estado = "style='margin-left: 10%; margin-right: 29%; float: left; position: relative; width: 11%; text-align: center; color:" + fondo + "; font-size: 14;'"
            If pasos_avanzados = 1 Then
                span_estado2 = "style='margin-left: 29%; float: left; position: relative; width: 11%; text-align: center; color:" + fondo + "; font-size: 14;'"
            End If
        End If

        avance_marca = (100 / (total_pasos - 1)) * pasos_avanzados
        divs_pasos = "<div id='divpasos' runat='server' class='status-rectangle' style='" + ancho + "'></div>"
        'Fin

        If codigo_estado_entrega <> 6 Then
            pasos_avanzados += 1
        End If

        Estilo = "<style type='text/css'> .delivery-status-text {display: inline-block; font-family: 'HelveticaNeueW01',Helvetica,Arial,sans-serif; font-weight: 400; font-style: normal; color: #336; font-size: 30px; "
        Estilo += "line-height: 1; vertical-align: middle; margin-left: 4%;} .status-orange {color: #f98c31;} .status-green {color: #218748;} .tracking-graphic-column {background-color: #f7f7f7; margin-bottom: 40px;} "
        Estilo += ".tracking-graphic-column .track-bar-holder {display: inline-block; vertical-align: middle; border: 4px solid #fff; border-radius: 22px; width: 70%; max-width: 670px; position: relative;} "
        Estilo += ".tracking-graphic-column .track-status-bar {overflow: hidden; border-radius: 22px; position: relative;} .tracking-graphic-column .track-status-bar .status-rectangle {float: left; height: 32px; "
        Estilo += "background-color: #d8d8d8; width: 20%; position: relative;} .tracking-graphic-column .track-status-bar .status-rectangle:nth-child(-n+" + pasos_avanzados.ToString + ") {background-color: " + fondo + ";} "
        Estilo += ".tracking-graphic-column .track-status-bar .status-rectangle:nth-child(-n+" + pasos_avanzados.ToString + "):before {border-left-color: " + fondo + ";} "
        Estilo += ".tracking-graphic-column .track-status-bar .status-rectangle:after, .tracking-graphic-column .track-status-bar .status-rectangle:before {width: 0; height: 0; border: 32px solid transparent; "
        Estilo += "border-left-width: 19.2px; border-right-width: 19.2px; border-left-color: white; content: ' '; position: absolute; right: -32px; top: -16px; z-index: 1;} "
        Estilo += ".tracking-graphic-column .track-status-bar .status-rectangle:before {right: -28.16px; z-index: 2; border-left-color: #d8d8d8;} "
        Estilo += ".tracking-graphic-column .track-bar-holder .tracking-circle {height: 32px; width: 32px; border: solid 3.2px #fff; box-shadow: 1px 1px 5px rgba(0,0,0,0.5); background-color: " + fondo + "; position: absolute; "
        Estilo += "display: block; left: calc(2% - 20px); border-radius: 50%; box-sizing: content-box; z-index: 5; top: -3.2px;} "
        Estilo += ".tracking-graphic-column .track-bar-holder .tracking-circle:after {content: ' '; height: 30%; width: 30%; left: 35%; top: 35%; background-color: #fff; z-index: 3; position: absolute; border-radius: 50%;} "
        Estilo += ".tracking-graphic-column .track-bar-holder .tracking-checkmark {height: 32px; width: 32px; border: solid 3.2px #fff; box-shadow: 1px 1px 5px rgba(0,0,0,0.5); background-color: " + fondo + "; "
        Estilo += "position: absolute; display: block; left: calc(2% - 20px); border-radius: 50%; "
        Estilo += "background-image: url('data:image/svg+xml;charset=US-ASCII,%3C%3Fxml%20version%3D%221.0%22%3F%3E%0A%3Csvg%20width%3D%2232%22%20height%3D%2224%22%20xmlns%3D"
        Estilo += "%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20xmlns%3Asvg%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%3E%0A%20%3Cg%3E%0A%20%20%3Ctitle%3ELayer%201%3C%2Ftit"
        Estilo += "le%3E%0A%20%20%3Cpath%20id%3D%22svg%5f2%22%20d%3D%22m6.956991%2C10.545466l1.906455%2C-1.996493l5.222733%2C5.465852l8.706619%2C-9.109631l1.906467%2C1.994915"
        Estilo += "l-10.613087%2C11.107201%22%20stroke-linecap%3D%22null%22%20stroke-linejoin%3D%2"
        Estilo += "2null%22%20stroke-dasharray%3D%22null%22%20stroke%3D%22%23ffffff%22%20fill%3D%22%23ffffff%22%2F%3E%0A%20%3C%2Fg%3E%0A%3C%2Fsvg%3E'); "
        Estilo += "box-sizing: content-box; z-index: 5; top: -3.2px; background-repeat: no-repeat; background-position: 0 4px;}"
        Estilo += ".tracking-graphic-column .track-bar-holder .tracking-checkmark:after {height: 30%; width: 30%; left: 35%; top: 35%; background-color: #fff; z-index: 3; position: absolute; border-radius: 50%;} </style>"

        If codigo_estado_entrega <> 6 Then
            pasos_avanzados -= 1
        End If

        Contenido = "<table><tr><td align='center'><p align='center' colspan='2'><b> Número de orden: " + codigoventa.ToString + "</b> </p><table><tr><td><p><b><font color='" + fondo + "'>" + Titulo + "</font></b></p>"
        Contenido += "</td></tr></table></br><table width='90%' border = '1'><tr><td colspan='2' align='center'><div id='DivTrackingRastreo' class='tracking-graphic-column' "
        Contenido += "style='margin-bottom: 0px; border: 3px solid grey;'><div style='margin-top: 3%; margin-bottom: 1%'><span id='SpanStatusRastreo' class='" + clase_spanrastreo + "' "
        Contenido += "style='width:40%'>" + texto_status + "</span></div><div class='track-bar-holder'><div id='DivTrackStatusBar' class='track-status-bar'>"
        'cantidad de divs de la barra
        For i As Integer = 1 To total_pasos - 1
            Contenido += divs_pasos.Replace("divpasos", "divpasos" + i.ToString)
        Next
        'Fin cantidad de divs
        Contenido += "</div><div id='DivMarcaRastreo' " + div_visible + " class='" + clase_marcarastreo + "' style='left: calc(" + avance_marca.ToString + "% - 20px);'></div></div>"
        Contenido += "<div id='DivEstados' style='overflow: hidden; margin-top: 1%; margin-bottom: 3%;'>"
        'Los estados de la entrega
        If total_pasos = 2 Then
            Contenido += "<span id='SpanEstado1' runat='server' " + span_estado + ">" + cadenas.GetValue(3).ToString + "</span>"
            If pasos_avanzados = 1 Then
                Contenido += "<span id='SpanEstado1' runat='server' " + span_estado2 + ">" + cadenas.GetValue(5).ToString + "</span>"
            End If
        Else
            For x As Integer = 1 To pasos_avanzados + 1
                If total_pasos = 3 Then
                    indice = x + 2
                ElseIf total_pasos = 6 Then
                    indice = x - 1
                End If
                Contenido += "<span id='SpanEstado" + x.ToString + "' runat='server' " + span_estado + ">" + cadenas(indice).ToString + "</span>"
            Next
        End If
        'Fin cantidad de estados
        'Datos de la venta
        Contenido += "</div></div></td></tr><tr><td align='left' width='50%'><font color='green'> Número de orden: </font> <b> " + codigoventa.ToString + " </b> <br/></td><td align='left' width='50%'><font color='green'> Fecha: </font> <b>" + FechaConfirmacion.ToString + "</b> <br/></td></tr>"
        Contenido += "<tr><td align='left' colspan='2'><font color='green'> Producto: </font> <b> " + NombreProducto.ToString + " </b> <br/></td></tr>"
        Contenido += "<tr><td align='left' width='50%'><font color='green'> Solicitante: </font> <b> " + NombreCliente.ToString + " </b> <br/></td><td align='left' width='50%'><font color='green'> Correo: </font><b>" + CorreoCliente.ToString + "</b><br/></td></tr>"
        Contenido += "<tr><td align='left' width='50%'><font color='green'> Telefono: </font> <b>" + Telefono1.ToString + "</b> <br/></td><td align='left' width='50%'><font color='green'> Telefono alterno: </font><b>" + Telefono2.ToString + "</b><br/></td></tr>"
        Contenido += "<tr><td align='left' colspan='2'><font color='green'> Direccion de entrega: </font><b>" + DireccionDeEntrega.ToString + "</b> <br/></td></tr>"

        If (CInt(CodigoFormaPago) <> 5) And (Cuotas = 1) Then
            Contenido += "<tr><td align='left' width='50%'><table><tr><td><font color='green'>Forma de pago: </font><b>" + FormaPago.ToString + "</b></td></tr></table></td>"
            Contenido += "<td align='left' width='50%'><table>"
            Contenido += "<tr><td><font color='green'>Precio: </font></td><td><b>" + Monto.ToString + "</b></td></tr>"
            If CDbl(Envio) > 0.00 Then
                Contenido += "<tr><td><font color='green'>Envío: </font></td><td><b>" + Envio.ToString + "</b></td></tr>"
            End If
            'Codigo Modificado por diego
            If IsNumeric(ServicioPago) Then
                If CDbl(ServicioPago) > 0 Then
                    Contenido += "<tr><td><font color='green'>Servicio pago en " + FormaPago.ToString + ": </font></td><td><b>" + ServicioPago.ToString + "</b></td></tr>"
                End If
            End If
            'Fin Codigo modificado por Diego
            Contenido += "<tr><td><font color='green'>Monto a pagar: </font></td><td><b>Q " + Total.ToString + "</b></td></tr></table>"
        End If
        If EsPedido = 1 Then
            Contenido += "<tr><td align='left' colspan='2'><font color='green'> El producto se entregará </font><b> " + DiasAereo.ToString + " días hábiles después de confirmar el pedido </b><br/></td></tr>"
        End If
        If (CInt(CodigoFormaPago) <> 5) And (Cuotas = 1) Then
            Contenido += "</td></tr>"
        End If
        Contenido += "<tr><td align='center' colspan='2'><b>Datos de la factura: </b><br/></td></tr><tr><td align='left' width='50%'><font color='green'> Nombre: </font><b>" + NombreFactura.ToString + "</b><br/></td>"
        Contenido += "<td align='left' width='50%'><font color='green'>NIT: </font><b>" + Nit.ToString + "</b></td></tr>"
        Contenido += "<tr><td align='left' colspan='2'><font color='green'> Dirección en factura: </font><b>" + DireccionFactura.ToString + "</b></td></tr>"
        Contenido += "<tr><td align='center' colspan='2'> <a href='" + PaginaProducto.ToString + "'>"
        Contenido += "<img src='" + Foto.ToString + "' /></a></td></tr>"
        If NombreRecibido.Trim.Length > 0 AndAlso FechaEntrega.Trim.Length > 0 Then
            Contenido += "<tr><td colspan='2'><table style='width: 100%'><tr><td align='left' colspan='1'><font color='green'> Entrega: </font><b>" + EmpresaDeEntrega.ToString + "</b> <br/></td></tr>"
            Contenido += "<tr><td align='left' colspan='1'><font color='green'> Recibe: </font> <b>" + NombreRecibido.ToString + "</b></td>"
            Contenido += "<td align='left' colspan='1'><font color='green'> Fecha entrega: </font> <b>" + FechaEntrega.ToString + "</b></td></tr></table></td></tr>"
        End If
        Contenido += "</table></br>"
        Contenido += "<table><tr><td align='left'>En cualquier momento puedes consultar el estado de tus ordenes, solo ingresa a la opción 'Mis compras' de la página, o haz "
        Contenido += "click aqui: <a href='http://www.guatemaladigital.com/ComprasRealizadas.aspx'>http://www.guatemaladigital.com/ComprasRealizadas.aspx</a> . </td></tr>"
        If guia.Trim.Length > 0 Then
            Contenido += "<tr><td align='left'>A partir del día de mañana puedes rastrear tu paquete haciendo click en el enlace Rastrear que se encuentra en la parte de arriba. </td></tr>"
        End If
        Contenido += "<tr><td align='left'> ¡Gracias por preferir GuatemalaDigital.com, como siempre haremos nuestro mejor esfuerzo para que quedes satisfecho con tu orden! </td></tr></table></br>"
        Contenido += "<table style='width:700px'><tr align='center'><td><img src='http://www.guatemaladigital.com/images/correopie.jpg'/></td></tr><tr align='center'><td>"
        Contenido += "</td></tr></table>"
        Contenido = Estilo + Contenido

        Try
            Enviar_Correo("info@guatemaladigital.com", CorreoCliente, Titulo, Contenido, "")
        Catch ex As Exception

        End Try

    End Sub

    'Sub Envio_SMS(ByVal telefono As String, ByVal mensaje As String)
    '    Dim snsclient As New AmazonSimpleNotificationServiceClient
    '    Dim snsrequest As New PublishRequest
    '    Dim snsresponse As New PublishResponse
    '    Dim snsmessagesettings As New Dictionary(Of String, MessageAttributeValue)()
    '    snsmessagesettings.Add("AWS.SNS.SMS.SenderID", New MessageAttributeValue() With {.StringValue = "GuatemalaDi", .DataType = "String"})
    '    snsmessagesettings.Add("AWS.SNS.SMS.SMSType", New MessageAttributeValue() With {.StringValue = "Transactional", .DataType = "String"})
    '    snsrequest.MessageAttributes = snsmessagesettings

    '    If telefono.Trim.Length < 11 Then
    '        If telefono.Trim.Length = 8 Then
    '            telefono = "502" + telefono.Trim
    '        Else
    '            Exit Sub
    '        End If
    '    End If

    '    If mensaje.Trim.Length > 160 Then
    '        mensaje = mensaje.Trim.Substring(0, 160)
    '    End If

    '    If telefono.Trim.Length = 11 AndAlso mensaje.Trim.Length > 10 Then
    '        snsrequest.PhoneNumber = telefono
    '        snsrequest.Message = mensaje
    '        snsresponse = snsclient.Publish(snsrequest)
    '    End If

    '    'Dim respuesta As String
    '    'respuesta = snsresponse.MessageId.ToString
    'End Sub


End Class


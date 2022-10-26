Imports System
Imports System.Net
Imports System.Web
Imports System.Text

Module Module1

    Dim mostrar As New Cargar
    Dim enviar As New Envio_De_Correos
    Dim MyConString As String
    Dim Consulta As String
    Dim Tablita As New DataTable
    Dim query As New StringBuilder

    Sub Main()
        MyConString = mostrar.Obtener_Cadena_Conexion
        Generar_CorreoExencion()
    End Sub

    Sub Generar_CorreoExencion()
        Dim Subject, Contenido As String
        Dim OrdenDeAmazon, ListaGuias, ListaImpuestos, ArchivosAdjuntos, NombreArchivo As String
        Dim TotalImpuesto As Decimal
        Dim CorreoAdministracion1, Correo, CorreoCopia As String
        Dim texto(), vectorImpuestos() As String
        Dim Exito As Boolean
        Dim Consulta2 As String
        Dim Mensaje As String
        Dim ErrorEnviarCorreo, EnviarCopiaExencion As Boolean
        Dim NumSolicitadoTransito, NumSinCancelar As Integer
        Dim CanalMostrarAlerta As String

        Mensaje = ""
        Correo = ""
        CorreoCopia = ""
        ErrorEnviarCorreo = False
        OrdenDeAmazon = ""
        CanalMostrarAlerta = ""

        CorreoAdministracion1 = mostrar.Obtener_Parametros("4", MyConString)

        Try
            CanalMostrarAlerta = ""
            query.Clear()
            query.AppendLine(" SELECT")
            query.AppendLine("   w.Url,")
            query.AppendLine("   ISNULL(Activo, 0) Activo")
            query.AppendLine(" FROM")
            query.AppendLine("   Alerta a,")
            query.AppendLine("   Webhook w")
            query.AppendLine(" WHERE")
            query.AppendLine("   a.CodigoWebhook = w.CodigoWebhook")
            query.AppendLine("   AND a.CodigoAlerta = 11")

            Tablita.Clear()
            mostrar.ejecuta_query_dt(query.ToString, Tablita, MyConString)

            For Each row As DataRow In Tablita.Rows
                If Not CBool(row("Activo").ToString) Then Continue For

                If row("Url").ToString.Equals("") Then Continue For

                CanalMostrarAlerta = row("Url").ToString
            Next

            Tablita.Clear()
            query.Clear()
            query.AppendLine(" SELECT")
            query.AppendLine("   OrdenDeAmazon")
            query.AppendLine(" FROM")
            query.AppendLine("   OrdenDeCompra")
            query.AppendLine(" WHERE")
            query.AppendLine("   CorreoDeExencion IS NULL")

            mostrar.ejecuta_query_dt(query.ToString, Tablita, MyConString)

            For Each Row As DataRow In Tablita.Rows
                OrdenDeAmazon = Row("OrdenDeAmazon").ToString

                If OrdenDeAmazon = "" Then Continue For

                ListaGuias = ""
                ListaImpuestos = ""
                TotalImpuesto = 0
                ArchivosAdjuntos = ""
                Mensaje = ""

                'verifica si hay líneas de pedido en solicitado, tránsito o pendiente, si hay no genera guía ni envía el correo
                query.Clear()
                query.AppendLine(" SELECT")
                query.AppendLine("   COUNT(1)")
                query.AppendLine(" FROM")
                query.AppendLine("   Pedido")
                query.AppendLine(" WHERE")
                query.Append("   OrdenDeAmazon = '").Append(OrdenDeAmazon).Append("'")
                query.AppendLine("   AND CodigoEstadoPedido IN (1, 2, 6)")
                query.AppendLine("   AND Cantidad > 0")

                NumSolicitadoTransito = mostrar.retornarentero(query.ToString, MyConString)

                If NumSolicitadoTransito <> 0 Then Continue For

                'obtiene el impuesto a pagar de la órden de amazon
                query.Clear()
                query.AppendLine(" SELECT")
                query.AppendLine("   ISNULL(SUM(Impuesto), 0) Impuesto")
                query.AppendLine(" FROM")
                query.AppendLine("   Pedido")
                query.AppendLine(" WHERE")
                query.Append("   OrdenDeAmazon = '").Append(OrdenDeAmazon).Append("'")
                query.AppendLine("   AND CodigoDeRastreo IS NOT NULL")

                TotalImpuesto = mostrar.retornardecimal(query.ToString, MyConString)

                query.Clear()
                query.AppendLine(" SELECT")
                query.AppendLine("   COUNT(1)")
                query.AppendLine(" FROM")
                query.AppendLine("   Pedido")
                query.AppendLine(" WHERE")
                query.Append("   OrdenDeAmazon = '").Append(OrdenDeAmazon).Append("'")
                query.AppendLine("   AND Cantidad > 0")
                query.AppendLine("   AND CodigoEstadoPedido <> 4")

                NumSinCancelar = mostrar.retornarentero(query.ToString, MyConString)

                'si el correo tiene SiempreEnviarExencion = 1, aunque totalimpuesto = 0 se envía el correo
                Consulta = "select COUNT(1) from CuentaDeCompra where Correo in ( select distinct Correo from Pedido where OrdenDeAmazon = '" & OrdenDeAmazon & "' and Correo is not null ) and SiempreEnviarExencion = 1 "
                query.Clear()
                query.AppendLine(" SELECT")
                query.AppendLine("   COUNT(1)")
                query.AppendLine(" FROM")
                query.AppendLine("   CuentaDeCompra")
                query.AppendLine(" WHERE")
                query.AppendLine("   Correo IN (")
                query.AppendLine("     SELECT")
                query.AppendLine("       DISTINCT Correo")
                query.AppendLine("     FROM")
                query.AppendLine("       Pedido")
                query.AppendLine("     WHERE")
                query.Append("       OrdenDeAmazon = '").Append(OrdenDeAmazon).Append("'")
                query.AppendLine("       AND Correo IS NOT NULL")
                query.AppendLine("   )")
                query.AppendLine("   AND SiempreEnviarExencion = 1")

                'si el correo tiene NuncaEnviarExencion = 1, aunque totalimpuesto > 0 no se envía el correo
                Consulta2 = "select COUNT(1) from CuentaDeCompra where Correo in ( " &
                                    "select distinct Correo from Pedido where OrdenDeAmazon = '" & OrdenDeAmazon & "' and Correo is not null " &
                                    ") and NuncaEnviarExencion = 1 "

                If Not (((TotalImpuesto > 0 Or mostrar.retornarentero(query.ToString, MyConString) > 0)) And mostrar.retornarentero(Consulta2, MyConString) = 0 And NumSinCancelar > 0) Then Continue For

                Generar_Guias(OrdenDeAmazon)
                
                If VerificarPaquete(OrdenDeAmazon, ListaGuias, ListaImpuestos) = False Then Continue For

                Exito = True
                texto = Split(ListaGuias, ",")
                vectorImpuestos = Split(ListaImpuestos, ",")

                For i As Integer = 0 To texto.Length - 1
                    'verifica si hay alguna línea en el tracking que no haya sido cancelada y tengan cantidad, si todas las líneas del tracking están canceladas no genera el pdf de ese tracking
                    query.Clear()
                    query.AppendLine(" SELECT")
                    query.AppendLine("   COUNT(1)")
                    query.AppendLine(" FROM")
                    query.AppendLine("   Pedido pe,")
                    query.AppendLine("   Paquete pa")
                    query.AppendLine(" WHERE")
                    query.AppendLine("   pe.CodigoPaquete = pa.CodigoPaquete")
                    query.AppendLine("   AND GuiaAerea = '" & texto(i) & "'")
                    query.AppendLine("   AND Cantidad > 0")
                    query.AppendLine("   AND CodigoEstadoPedido <> 4")

                    If mostrar.retornarentero(query.ToString, MyConString) = 0 Then Continue For 'genera el pdf

                    If Exito = False Then Continue For

                    'Revisar

                    NombreArchivo = ""
                    Dim FormGuia As New GuiaPdf
                    Exito = FormGuia.Crear_Reporte_Pdf(texto(i), NombreArchivo, MyConString, Mensaje)

                    If NombreArchivo = "" Then Continue For

                    If ArchivosAdjuntos = "" Then
                        ArchivosAdjuntos = NombreArchivo
                    Else
                        If InStr(ArchivosAdjuntos, NombreArchivo) = 0 Then
                            ArchivosAdjuntos = ArchivosAdjuntos + "," + NombreArchivo
                        End If
                    End If
                Next

                If Exito = False Then
                    Subject = "Error procedimiento Correo de Exencion"
                    Contenido = "Orden Amazon: " + OrdenDeAmazon + ", " + Mensaje
                    
                    If CanalMostrarAlerta <> "" Then
                        EnviarAlertaSlackPromocion("ALERTA-EXENCION", Contenido, Subject, "Generar_CorreoExencion", CanalMostrarAlerta)
                    End If

                    Return
                End If

                'Just here

                query.Clear()
                query.AppendLine(" SELECT")
                query.AppendLine("   COUNT(1)")
                query.AppendLine(" FROM")
                query.AppendLine("   Pedido")
                query.AppendLine(" WHERE")
                query.AppendLine("   OrdenDeAmazon = '" & OrdenDeAmazon & "'")
                query.AppendLine("   AND Correo IS NOT NULL")

                If mostrar.retornarentero(query.ToString, MyConString) = 0 Then Continue For

                Correo = ""
                query.Clear()
                query.AppendLine(" SELECT")
                query.AppendLine("   TOP(1) ISNULL(pe.Correo, '') Correo,")
                query.AppendLine("   ISNULL(cc.EnviarCopiaExencion, 0) EnviarCopiaExencion")
                query.AppendLine(" FROM")
                query.AppendLine("   Pedido pe,")
                query.AppendLine("   CuentaDeCompra cc")
                query.AppendLine(" WHERE")
                query.AppendLine("   pe.Correo = cc.Correo")
                query.AppendLine("   AND OrdenDeAmazon = '" & OrdenDeAmazon & "'")
                query.AppendLine("   AND pe.Correo IS NOT NULL")

                Tablita.Clear()
                mostrar.ejecuta_query_dt(query.ToString, Tablita, MyConString)

                For Each rowcita As DataRow In Tablita.Rows
                    Correo = rowcita("Correo").ToString
                    EnviarCopiaExencion = CBool(rowcita("EnviarCopiaExencion").ToString)

                    If EnviarCopiaExencion = True Then
                        CorreoCopia = CorreoAdministracion1
                    Else
                        CorreoCopia = ""
                    End If
                Next

                'enviar correo
                Subject = "Tax exempt for " & Correo

                Contenido = "Good day, attached proof of export.<br/><br/>E-mail account: " & Correo & "<br/>Order numbers:<br/>" & OrdenDeAmazon & "<br/><br/><br/>Best Regards,<br/>Mario Porres<br/><br/>"

                If ArchivosAdjuntos = "" Then Continue For

                If enviar.Enviar_Correo_Con_Attachment_Verificar_Error(Correo, "tax-exempt@amazon.com", Subject, Contenido, CorreoCopia, ArchivosAdjuntos) = True Then
                    'actualizar campo CorreoDeExencion
                    Consulta = "Update OrdenDeCompra set CorreoDeExencion = GETDATE() Where OrdenDeAmazon = '" & OrdenDeAmazon & "'"
                    mostrar.insertarmodificareliminar(Consulta, MyConString)
                    Return
                End If

                Subject = "Error al enviar el correo en el procedimiento Correo Exencion "
                Contenido = "Error procedimiento Correo Exencion al enviar el correo desde: " + Correo + ", Orden de Amazon: " + OrdenDeAmazon + "<br/><br/>" + Contenido

                If CanalMostrarAlerta <> "" Then
                    EnviarAlertaSlackPromocion("ALERTA-EXENCION", Contenido, Subject, "Generar_CorreoExencion", CanalMostrarAlerta)
                End If
            Next Row
        Catch ex As Exception
            Consulta = "select url from Webhook where CodigoWebhook = 1" 'canal de alerta sistemas
            CanalMostrarAlerta = mostrar.retornarcadena(Consulta, MyConString)
            Subject = "Error procedimiento Correo de Exencion"
            Contenido = "Orden Amazon: " + OrdenDeAmazon + Chr(13) + ex.Message.ToString
            EnviarAlertaSlackPromocion("ALERTA-EXENCION", Contenido, Subject, "Generar_CorreoExencion", CanalMostrarAlerta)
        End Try
    End Sub

    'Gloria 07-mar-2018, Tarea 245
    Sub EnviarAlertaSlackPromocion(ByVal TituloAlerta As String, ByVal Mensaje As String, ByVal Subject As String, ByVal Procedimiento As String, ByVal UrlCanalSlack As String)
        Dim jSonSlack As String = ""

        Mensaje = Replace(Mensaje, """", "")

        jSonSlack += "{ "
        jSonSlack += "    ""text"": """ & TituloAlerta & """, "
        jSonSlack += "    ""attachments"": [ "
        jSonSlack += "        { "
        jSonSlack += "            ""title"": """ + Subject + """, "
        jSonSlack += "            ""fields"": [ "
        jSonSlack += "                { "
        jSonSlack += "                    ""title"": """ + "" + """, "
        jSonSlack += "                    ""value"": """ + Mensaje + """, "
        jSonSlack += "                    ""Short"": ""True"" "
        jSonSlack += "                } "
        jSonSlack += "            ], "
        jSonSlack += "            ""author_name"": """ + Procedimiento + "GuatemalaDigital.com/"", "
        jSonSlack += "            ""author_icon"": ""https://guatemaladigital.com/images/favicon.ico"" "
        jSonSlack += "        } "
        jSonSlack += "    ] "
        jSonSlack += "} "

        QuitaAcentos(jSonSlack)
        Enviar_Resultados_Slack(jSonSlack, UrlCanalSlack)
    End Sub

    Public Sub Enviar_Resultados_Slack(json As String, ByVal UrlCanalSlack As String)
        Dim Cadena As String = ""

        Try
            Dim webClient = New WebClient()
            webClient.Headers(HttpRequestHeader.ContentType) = "application/json; chartset=utf-8"
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12
            webClient.UploadString(UrlCanalSlack, json)
        Catch ex As Exception
            Cadena = "Error al enviar mensaje Slack :  " + "<br/><br/>" + ex.Message.ToString + "<br/><br/>"
        End Try
    End Sub

    'Gloria Tarea 204,205
    Public Sub QuitaAcentos(ByRef msj As String)
        msj = Replace(msj, "á", "a")
        msj = Replace(msj, "é", "e")
        msj = Replace(msj, "í", "i")
        msj = Replace(msj, "ó", "o")
        msj = Replace(msj, "ú", "u")
        msj = Replace(msj, "Á", "A")
        msj = Replace(msj, "É", "E")
        msj = Replace(msj, "Í", "I")
        msj = Replace(msj, "Ó", "O")
        msj = Replace(msj, "Ú", "U")
        msj = Replace(msj, "\", "/")
        msj = Replace(msj, "'", " ")
    End Sub

    Sub Generar_Guias(ByVal OrdenDeAmazon As String)
        ' buscar las ordenes de amazon en la tabla ordendecompra que no se haya enviado correo de excepcion a Amazon
        ' busca las ordenes que no tengan líneas de pedido con tracking nulo 

        'verifica que haya algún producto recibido en la orden
        query.Clear()
        query.AppendLine(" SELECT")
        query.AppendLine("   COUNT(1)")
        query.AppendLine(" FROM")
        query.AppendLine("   Pedido")
        query.AppendLine(" WHERE")
        query.AppendLine("   OrdenDeAmazon = '" & OrdenDeAmazon & "'")
        query.AppendLine("   AND CodigoEstadoPedido = 3")
        query.AppendLine("   AND Cantidad > 0")

        If mostrar.retornarentero(query.ToString, MyConString) > 0 Then 'la orden tienen algún producto recibido
            Generar_Guia_Insertar_Paquetes(OrdenDeAmazon)
        End If
    End Sub

    'función que verifica si todos los tracking de la orden de amazon en la tabla pedido también se encuentran en la tabla paquete
    'si todos los tracking se encuentra, devuelve las guías y el impuesto pagado de todos los tracking
    'también devuelve el total del impuesto pagado de todos los tracking
    'si no se encuentran todos los tracking o si alguna línea de pedido no tiene tracking, devolverá false
    Function VerificarPaquete(ByVal OrdenDeAmazon As String, ByRef ListaGuias As String, ByRef ListaImpuestos As String) As Boolean
        Dim TotalPedido, TotalPaquete As Integer
        Dim Exito As Boolean
        Dim CodigoDeRastreo As String
        'Dim Consulta As String
        'Dim TrackingNulo As Integer
        Dim GuiaAerea As String

        'número de líneas de la orden que no tienen tracking (codigoderastreo = null)
        'query.Clear()
        'query.AppendLine(" SELECT")
        'query.AppendLine("   COUNT(1)")
        'query.AppendLine(" FROM")
        'query.AppendLine("   Pedido")
        'query.AppendLine(" WHERE")
        'query.AppendLine("   OrdenDeAmazon = '" & OrdenDeAmazon & "'")
        'query.AppendLine("   AND (")
        'query.AppendLine("     CodigoDeRastreo IS NULL")
        'query.AppendLine("     OR LEN(codigoderastreo) = 0")
        'query.AppendLine("   )")

        'TrackingNulo = mostrar.retornarentero(query.ToString, MyConString)

        'total de tracking de la orden
        query.Clear()
        query.AppendLine(" SELECT")
        query.AppendLine("   COUNT(1)")
        query.AppendLine(" FROM")
        query.AppendLine(" (")
        query.AppendLine("   SELECT")
        query.AppendLine("     DISTINCT CodigoDeRastreo")
        query.AppendLine("   FROM")
        query.AppendLine("     Pedido")
        query.AppendLine("   WHERE")
        query.AppendLine("     OrdenDeAmazon = '" & OrdenDeAmazon & "'")
        query.AppendLine("     AND CodigoDeRastreo IS NOT NULL")
        query.AppendLine("     AND CodigoEstadoPedido <> 4")
        query.AppendLine(" ) z")

        TotalPedido = mostrar.retornarentero(query.ToString, MyConString)
        TotalPaquete = 0
        Exito = False
        ListaGuias = ""
        ListaImpuestos = ""

        If TotalPedido = 0 Then Return Exito

        Tablita.Clear()
        query.Clear()
        query.AppendLine(" SELECT")
        query.AppendLine("   CodigoDeRastreo,")
        query.AppendLine("   CASE")
        query.AppendLine("     WHEN CHARINDEX('(', CodigoDeRastreo) > 0")
        query.AppendLine("     AND CHARINDEX(')', CodigoDeRastreo) > 0")
        query.AppendLine("     THEN SUBSTRING(")
        query.AppendLine("       CodigoDeRastreo,")
        query.AppendLine("       CHARINDEX('(', CodigoDeRastreo) + 1,")
        query.AppendLine("       CHARINDEX(')', CodigoDeRastreo) - CHARINDEX('(', CodigoDeRastreo) - 1")
        query.AppendLine("     )")
        query.AppendLine("     ELSE CodigoDeRastreo")
        query.AppendLine("   END Tracking,")
        query.AppendLine("   ISNULL(SUM(Impuesto), 0) Impuesto")
        query.AppendLine(" FROM")
        query.AppendLine("   Pedido")
        query.AppendLine(" WHERE")
        query.AppendLine("   OrdenDeAmazon = '" & OrdenDeAmazon & "'")
        query.AppendLine("   AND CodigoEstadoPedido <> 4")
        query.AppendLine(" GROUP BY")
        query.AppendLine("   CodigoDeRastreo")
        query.AppendLine(" HAVING")
        query.AppendLine("   CodigoDeRastreo IS NOT NULL")

        mostrar.ejecuta_query_dt(query.ToString, Tablita, MyConString)

        For Each row As DataRow In Tablita.Rows
            If Not row("Impuesto").ToString.Equals("") Then
                If ListaImpuestos = "" Then
                    ListaImpuestos = row("Impuesto").ToString
                Else
                    ListaImpuestos += "," + row("Impuesto").ToString
                End If
            End If

            If row("Tracking").ToString.Equals("") Then Continue For

            CodigoDeRastreo = row("Tracking").ToString

            Consulta = "select COUNT(1) from Paquete where codigoderastreo like '%" & CodigoDeRastreo & "%' "

            If mostrar.retornarentero(Consulta, MyConString) <> 1 Then Continue For

            TotalPaquete = TotalPaquete + 1
            Consulta = "select isnull(GuiaAerea,'') from Paquete where codigoderastreo like '%" & CodigoDeRastreo & "%' "
            GuiaAerea = mostrar.retornarcadena(Consulta, MyConString).Trim

            If GuiaAerea = "" Then
                GuiaAerea = Obtener_Nueva_Guia()
                Consulta = "update paquete set GuiaAerea = '" & GuiaAerea & "', Generado = 1 where codigoderastreo like '%" & CodigoDeRastreo & "%' "
                mostrar.insertarmodificareliminar(Consulta, MyConString)
            End If

            If ListaGuias = "" Then
                ListaGuias = GuiaAerea
            Else
                ListaGuias = ListaGuias & "," & GuiaAerea
            End If
        Next row

        If TotalPedido = TotalPaquete Then Exito = True

        Return Exito
    End Function

    'inserta los tracking faltantes de la orden de amazon a la tabla paquete
    Sub Generar_Guia_Insertar_Paquetes(ByVal OrdenDeAmazon As String)
        Dim Guia As String
        Dim CodigoDeRastreo As String
        Dim CadFechaOrden, Descripcion, EnviadoPor As String
        Dim Contador As Integer
        Dim Exito As Boolean
        Dim Peso, Impuesto As Decimal
        Dim CadPeso, CadImpuesto As String
        Dim CodigoPaquete As Integer
        Dim CodigoEstablecimiento As String

        'me devuelve el tracking sin la empresa de entrega, para UPS(1Z1020) devuelve 1Z1020
        query.Clear()
        query.AppendLine(" SELECT")
        query.AppendLine("   DISTINCT pe.CodigoDeRastreo,")
        query.AppendLine("   (")
        query.AppendLine("     SELECT")
        query.AppendLine("       COUNT(1)")
        query.AppendLine("     FROM")
        query.AppendLine("       Pedido")
        query.AppendLine("     WHERE")
        query.AppendLine("       CodigoEstadoPedido = 3")
        query.AppendLine("       AND CodigoDeRastreo = pe.CodigoDeRastreo")
        query.AppendLine("   ) Recibidos,")
        query.AppendLine("   CASE")
        query.AppendLine("     WHEN CHARINDEX('(', pe.CodigoDeRastreo) > 0")
        query.AppendLine("     AND CHARINDEX(')', pe.CodigoDeRastreo) > 0")
        query.AppendLine("     THEN SUBSTRING(")
        query.AppendLine("       pe.CodigoDeRastreo,")
        query.AppendLine("       CHARINDEX('(', pe.CodigoDeRastreo) + 1,")
        query.AppendLine("       CHARINDEX(')', pe.CodigoDeRastreo) - CHARINDEX('(', pe.CodigoDeRastreo) - 1")
        query.AppendLine("     )")
        query.AppendLine("     ELSE pe.CodigoDeRastreo")
        query.AppendLine("   END TrackingCorto")
        query.AppendLine(" FROM")
        query.AppendLine("   Pedido pe")
        query.AppendLine(" WHERE")
        query.AppendLine("   pe.OrdenDeAmazon = '" & OrdenDeAmazon & "'")
        query.AppendLine("   AND CodigoDeRastreo IS NOT NULL")
        query.AppendLine("   AND CodigoEstadoPedido <> 4")

        Tablita.Clear()
        mostrar.ejecuta_query_dt(query.ToString, Tablita, MyConString)

        For Each row As DataRow In Tablita.Rows
            If Not row("CodigoDeRastreo").ToString.Equals("") Then
                CodigoDeRastreo = row("CodigoDeRastreo")
            Else
                CodigoDeRastreo = ""
            End If

            'verifica que el tracking no exista en la tabla paquete
            Consulta = "Select count(1) from Paquete where CodigoDeRastreo like '%" & CodigoDeRastreo & "%'"
            If mostrar.retornarentero(Consulta, MyConString) <> 0 Then
                'verifica si existe registros en la tabla pedido con el código de rastreo y sin número de paquete (sin guía)

                Consulta = "select count(1) from Pedido where CodigoDeRastreo like '%" & CodigoDeRastreo & "%' and CodigoPaquete is null"
                If mostrar.retornarentero(Consulta, MyConString) = 0 Then Return

                'agrega el código de paquete al registro de pedido

                Consulta = "Select CodigoPaquete from Paquete where CodigoDeRastreo like '%" & CodigoDeRastreo & "%'"
                CodigoPaquete = mostrar.retornarentero(Consulta, MyConString)

                query.Clear()
                query.AppendLine(" UPDATE")
                query.AppendLine("   Pedido")
                query.AppendLine(" SET")
                query.AppendLine("   CodigoPaquete = " & CStr(CodigoPaquete) & "")
                query.AppendLine(" WHERE")
                query.AppendLine("   CodigoDeRastreo IN (")
                query.AppendLine("     SELECT")
                query.AppendLine("       TOP(1) CodigoDeRastreo")
                query.AppendLine("     FROM")
                query.AppendLine("       Pedido")
                query.AppendLine("     WHERE")
                query.AppendLine("       CodigoDeRastreo = '" & CodigoDeRastreo & "'")
                query.AppendLine("       OR CodigoDeRastreo LIKE '%" & CodigoDeRastreo & "%'")
                query.AppendLine("   )")

                mostrar.insertarmodificareliminar(query.ToString, MyConString)

                Return
            End If

            'verifica si el tracking existe en tabla pedido
            Consulta = "select count(1) from Pedido where CodigoDeRastreo like '%" & CodigoDeRastreo & "%'"

            If mostrar.retornarentero(Consulta, MyConString) = 0 Then Continue For

            'obtiene los datos que se guardarán en la tabla paquete
            Consulta = "select top(1) CONVERT(date,FechaDeOrden) from Pedido where CodigoDeRastreo like '%" & CodigoDeRastreo & "%'"
            CadFechaOrden = mostrar.retornafechaSinHora(Consulta, MyConString)
            CadFechaOrden = CadFechaOrden.Substring(6, 4) & "-" & CadFechaOrden.Substring(3, 2) & "-" & CadFechaOrden.Substring(0, 2)

            Consulta = "DECLARE @valores VARCHAR(1000) select @valores= COALESCE(@valores + ', ', '') + t.TipoDeProducto from ( select distinct(isnull(TipoDeProducto,'')) as TipoDeProducto from Pedido where CodigoDeRastreo like '%" & CodigoDeRastreo & "%' ) t select replace(@valores,',',' ') "

            Descripcion = mostrar.retornarcadena(Consulta, MyConString)

            Consulta = "select top(1) isnull(Vendedor,'') from Pedido where CodigoDeRastreo like '%" & CodigoDeRastreo & "%'"
            EnviadoPor = mostrar.retornarcadena(Consulta, MyConString)

            'verifica si hay productos sin peso
            'si hay algún producto sin peso (NULL), Peso = 0
            'si todos los productos tienen peso, se suman todos
            Consulta = "select count(1) from Producto where CodigoAmazon in ( " &
                               "select CodigoAmazon from Pedido where CodigoDeRastreo like '%" & CodigoDeRastreo & "%' " &
                                ") and Peso is null "

            If mostrar.retornarentero(Consulta, MyConString) = 0 Then
                Consulta = "select sum(isnull(Peso,0)) from Producto where CodigoAmazon in ( select CodigoAmazon from Pedido where CodigoDeRastreo like '%" & CodigoDeRastreo & "%' )"

                Peso = mostrar.retornardecimal(Consulta, MyConString)
            Else
                Peso = 0
            End If

            Consulta = "select sum(isnull(Impuesto,0)) from Pedido where CodigoDeRastreo like '%" & CodigoDeRastreo & "%'"
            Impuesto = mostrar.retornardecimal(Consulta, MyConString)

            CadPeso = ""
            If Peso <> 0 Then
                CadPeso = CStr(Peso)
            End If

            CadImpuesto = ""

            Guia = Obtener_Nueva_Guia()

            CodigoEstablecimiento = ""
            Consulta = "select top(1) pr.CodigoEstablecimiento from Pedido pe, Producto pr where pe.CodigoAmazon = pr.CodigoAmazon and pe.OrdenDeAmazon = '" & OrdenDeAmazon & "' and CodigoDeRastreo is not null "
            Tablita.Clear()
            mostrar.ejecuta_query_dt(Consulta, Tablita, MyConString)

            For Each rowcita As DataRow In Tablita.Rows
                If rowcita("CodigoEstablecimiento").ToString.Equals("") Then Continue For
                CodigoEstablecimiento = rowcita("CodigoEstablecimiento").ToString
            Next rowcita

            Contador = 0
            Exito = False

            Insertar_Paquete(CodigoDeRastreo, Guia, EnviadoPor, Descripcion, CadPeso, "", CadImpuesto, "1", "1", CodigoEstablecimiento, CodigoPaquete)
            Actualizar_Estado_Paquete_Venta(CodigoPaquete, CodigoDeRastreo, True)
        Next row
    End Sub

    'obtiene el valor para una guía generada por el sistema para un paquete que no tiene guía
    Function Obtener_Nueva_Guia() As String
        Dim Guia As String
        Dim Contador As Integer
        Dim Exito As Boolean

        Consulta = "declare @Guia as int " &
                    "select @Guia = CAST(valor AS int) from Parametro where CodigoParametro = 51 " &
                    "Update Parametro set Valor = CAST(@Guia + 1 as varchar) where CodigoParametro = 51 " &
                    "select @Guia "

        Guia = CStr(mostrar.retornarentero(Consulta, MyConString))

        Contador = 0
        Exito = False
        While Not Exito 'And Contador < 30
            Consulta = "Select count(1) from Paquete Where GuiaAerea = '" & "G" & Guia & "'"
            If mostrar.retornarentero(Consulta, MyConString) = 0 Then
                Exito = True
            Else
                Contador = Contador + 1
                Guia = CStr(CInt(Guia) + 1)
            End If
        End While

        If Contador > 0 Then
            Consulta = "Update Parametro set Valor = " & CStr(CInt(Guia) + 1) & " where CodigoParametro = 51 "
            mostrar.insertarmodificareliminar(Consulta, MyConString)
        End If

        Obtener_Nueva_Guia = "G" & Guia
    End Function


    Sub Insertar_Paquete(ByVal Tracking As String,
                         ByVal GuiaAerea As String,
                         ByVal EnviadoPor As String,
                         ByVal Descripcion As String,
                         ByVal Peso As String,
                         ByVal PesoVolumetrico As String,
                         ByVal MontoImpuesto As String,
                         ByVal Generado As String,
                         ByVal EstadoPaquete As String,
                         ByVal CodigoEstablecimiento As String,
                         ByRef CodigoPaquete As Integer)
        If GuiaAerea <> "" Then
            GuiaAerea = "'" & GuiaAerea & "'"
        Else
            GuiaAerea = "NULL"
        End If

        If EnviadoPor <> "" Then
            EnviadoPor = "'" & Replace(EnviadoPor, "'", "''") & "'"
        Else
            EnviadoPor = "NULL"
        End If

        If Descripcion <> "" Then
            Descripcion = "'" & Replace(Descripcion, "'", "''") & "'"
        Else
            Descripcion = "NULL"
        End If

        If Peso <> "" Then
            If IsNumeric(Peso) = False Then
                Peso = "NULL"
            End If
        Else
            Peso = "NULL"
        End If

        If PesoVolumetrico <> "" Then
            If IsNumeric(PesoVolumetrico) = False Then
                PesoVolumetrico = "NULL"
            End If
        Else
            PesoVolumetrico = "NULL"
        End If

        If MontoImpuesto <> "" Then
            If IsNumeric(MontoImpuesto) = False Then
                MontoImpuesto = "NULL"
            End If
        Else
            MontoImpuesto = "NULL"
        End If

        If EstadoPaquete = "" Then
            EstadoPaquete = "NULL"
        End If

        If CodigoEstablecimiento = "" Then
            CodigoEstablecimiento = "NULL"
        End If

        Consulta = "select count(1) from Paquete"
        If mostrar.retornarentero(Consulta, MyConString) > 0 Then
            Consulta = "select max(CodigoPaquete) + 1 from Paquete"
            CodigoPaquete = mostrar.retornarentero(Consulta, MyConString)
        Else
            CodigoPaquete = 1
        End If

        Consulta = "insert into Paquete (CodigoPaquete,CodigoDeRastreo, GuiaAerea, EnviadoPor, Descripcion, Peso, PesoVolumetrico, MontoImpuesto, Fecha, Generado, CodigoEstadoPaquete, CodigoEstablecimiento) " &
                    "values (" & CStr(CodigoPaquete) & ",'" & Tracking & "'," & GuiaAerea & "," & EnviadoPor & "," & Descripcion & "," & Peso & "," & PesoVolumetrico & "," & MontoImpuesto & ",getdate()," & Generado & "," & EstadoPaquete & "," & CodigoEstablecimiento & ")"

        mostrar.insertarmodificareliminar(Consulta, MyConString)
    End Sub

    Sub Actualizar_Estado_Paquete_Venta(ByVal CodigoPaquete As Integer,
                                        ByVal Tracking As String,
                                        ByVal WareHouseNotificacion As Boolean)
        Dim TrackingReducido, CadTracking, Mensaje, MensajeError As String
        Dim dtventas As New DataTable

        TrackingReducido = "" : Mensaje = "" : MensajeError = ""
        Validar_Tracking(Tracking, TrackingReducido, MensajeError)

        If TrackingReducido <> "" Then
            CadTracking = TrackingReducido
        Else
            CadTracking = Tracking
        End If

        If MensajeError = "Se encontró más de 1 tracking" Then Return

        query.Clear()
        query.AppendLine(" UPDATE")
        query.AppendLine("   Pedido")
        query.AppendLine(" SET")
        query.AppendLine("   CodigoPaquete = " & CStr(CodigoPaquete) & "")
        query.AppendLine(" WHERE")
        query.AppendLine("   CodigoDeRastreo IN (")
        query.AppendLine("     SELECT")
        query.AppendLine("       TOP(1) CodigoDeRastreo")
        query.AppendLine("     FROM")
        query.AppendLine("       Pedido")
        query.AppendLine("     WHERE")
        query.AppendLine("       CodigoDeRastreo = '" & CadTracking & "'")
        query.AppendLine("       OR CodigoDeRastreo LIKE '%" & CadTracking & "%'")
        query.AppendLine("   )")

        mostrar.insertarmodificareliminar(query.ToString, MyConString)

        'Actualizacion y envio de correo de estado de entrega
        Consulta = "Select distinct vp.CodigoVenta from VentaPedido vp inner join Pedido p On vp.CodigoPedido = p.CodigoPedido inner join Venta v On vp.CodigoVenta = v.CodigoVenta inner join Paquete pa On p.CodigoPaquete = pa.CodigoPaquete where v.CodigoEstadoEntrega = 2 And pa.GuiaAerea Is Not null And p.CodigoPaquete = " + CodigoPaquete.ToString.Trim
        mostrar.ejecuta_query_dt(Consulta, dtventas, MyConString)

        If dtventas.Rows.Count > 0 Then
            Try
                For Each fila As DataRow In dtventas.Rows
                    'Actualiza el codigoestadoentrega 3 en venta
                    Consulta = "update Venta Set CodigoEstadoEntrega = 3 where CodigoVenta = " + fila("CodigoVenta").ToString + " And CodigoEstadoEntrega = 2"
                    mostrar.insertarmodificareliminar(Consulta, MyConString)

                    'Envio de correo de estado de entrega
                    enviar.Enviar_Correo_Rastreo(fila("CodigoVenta").ToString, MyConString)
                Next
            Catch ex As Exception : End Try
        End If
        'Fin Actualizacion

        If WareHouseNotificacion Then
            Consulta = "update Pedido Set CodigoEstadoPedido = 2 where CodigoDeRastreo In (Select top(1) CodigoDeRastreo from Pedido where CodigoDeRastreo = '" & CadTracking & "' or CodigoDeRastreo like '%" & CadTracking & "%' ) and CodigoEstadoPedido = 1 "
            mostrar.insertarmodificareliminar(Consulta, MyConString)
        End If
    End Sub

    Function Validar_Tracking(ByVal Tracking As String, ByRef TrackingReducido As String, ByRef MensajeError As String) As Boolean
        Dim Exito As Boolean

        Exito = True
        MensajeError = ""
        TrackingReducido = ""

        If Len(Tracking) > 60 Then
            MensajeError = "Tracking demasiado largo"
            Return False
        End If

        'verifica si existe el tracking ingresado
        query.Clear()
        query.AppendLine(" SELECT")
        query.AppendLine("   COUNT(1)")
        query.AppendLine(" FROM")
        query.AppendLine("   Pedido")
        query.AppendLine(" WHERE")
        query.AppendLine("   (")
        query.AppendLine("     CodigoDeRastreo = '" & Tracking & "'")
        query.AppendLine("     OR CodigoDeRastreo LIKE '%" & Tracking & "%'")
        query.AppendLine("   )")
        query.AppendLine("   AND CodigoEstadoPedido <> 4")

        If mostrar.retornarentero(query.ToString, MyConString) <> 0 Then 'no hay tracking
            'verifica si se encontró más de 1 tracking
            query.Clear()
            query.AppendLine(" SELECT")
            query.AppendLine("   COUNT(1)")
            query.AppendLine(" FROM")
            query.AppendLine(" (")
            query.AppendLine("   SELECT")
            query.AppendLine("     DISTINCT REPLACE(")
            query.AppendLine("       SUBSTRING(")
            query.AppendLine("         codigoderastreo,")
            query.AppendLine("         (CHARINDEX('(', codigoderastreo) + 1),")
            query.AppendLine("         (")
            query.AppendLine("           LEN(codigoderastreo) - CHARINDEX('(', codigoderastreo)")
            query.AppendLine("         )")
            query.AppendLine("       ),")
            query.AppendLine("       ')',")
            query.AppendLine("       ''")
            query.AppendLine("     ) CodigoDeRastreo")
            query.AppendLine("   FROM")
            query.AppendLine("     Pedido")
            query.AppendLine("   WHERE")
            query.AppendLine("     (")
            query.AppendLine("       CodigoDeRastreo = '" & Tracking & "'")
            query.AppendLine("       OR CodigoDeRastreo LIKE '%" & Tracking & "%'")
            query.AppendLine("     )")
            query.AppendLine("     AND CodigoEstadoPedido <> 4")
            query.AppendLine(" ) t")

            If mostrar.retornarentero(query.ToString, MyConString) > 1 Then
                MensajeError = "Se encontró más de 1 tracking"
                Exito = False
            End If

            Return Exito
        End If

        'verifica si el tracking es númerico (puede ser de Fedex)
        If IsNumeric(Tracking) = False Then
            MensajeError = "El tracking no existe"
            Return False
        End If

        TrackingReducido = Microsoft.VisualBasic.Right(Tracking, 12)
        'verifica si existe el tracking 
        query.Clear()
        query.AppendLine(" SELECT")
        query.AppendLine("   COUNT(1)")
        query.AppendLine(" FROM")
        query.AppendLine("   Pedido")
        query.AppendLine(" WHERE")
        query.AppendLine("   (")
        query.AppendLine("     CodigoDeRastreo = '" & TrackingReducido & "'")
        query.AppendLine("     OR CodigoDeRastreo LIKE '%" & TrackingReducido & "%'")
        query.AppendLine("   )")
        query.AppendLine("   AND CodigoEstadoPedido <> 4")

        If mostrar.retornarentero(query.ToString, MyConString) = 0 Then 'no hay tracking
            MensajeError = "El tracking no existe"
            Exito = False
            Return Exito
        End If

        'verifica si se encontró más de 1 tracking
        query.Clear()
        query.AppendLine(" SELECT")
        query.AppendLine("   COUNT(1)")
        query.AppendLine(" FROM")
        query.AppendLine(" (")
        query.AppendLine("   SELECT")
        query.AppendLine("     DISTINCT REPLACE(")
        query.AppendLine("       SUBSTRING(")
        query.AppendLine("         codigoderastreo,")
        query.AppendLine("         (CHARINDEX('(', codigoderastreo) + 1),")
        query.AppendLine("         (")
        query.AppendLine("           LEN(codigoderastreo) - CHARINDEX('(', codigoderastreo)")
        query.AppendLine("         )")
        query.AppendLine("       ),")
        query.AppendLine("       ')',")
        query.AppendLine("       ''")
        query.AppendLine("     ) CodigoDeRastreo")
        query.AppendLine("   FROM")
        query.AppendLine("     Pedido")
        query.AppendLine("   WHERE")
        query.AppendLine("     (")
        query.AppendLine("       CodigoDeRastreo = '" & TrackingReducido & "'")
        query.AppendLine("       OR CodigoDeRastreo LIKE '%" & TrackingReducido & "%'")
        query.AppendLine("     )")
        query.AppendLine("     AND CodigoEstadoPedido <> 4")
        query.AppendLine(" ) t")

        If mostrar.retornarentero(query.ToString, MyConString) > 1 Then
            MensajeError = "Se encontró más de 1 tracking"
            Exito = False
        End If

        Return Exito
    End Function
End Module

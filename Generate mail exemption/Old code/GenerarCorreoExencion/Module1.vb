Imports System
Imports System.Net
Imports System.Web

Module Module1

    Dim mostrar As New Cargar
    Dim enviar As New Envio_De_Correos
    Dim MyConString As String
    Dim Consulta As String
    Dim Tablita As New DataTable


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
        Dim i As Integer
        Dim Exito As Boolean
        Dim Consulta2 As String
        Dim Mensaje As String
        Dim ErrorEnviarCorreo, EnviarCopiaExencion As Boolean
        Dim NumSolicitadoTransito, NumSinCancelar As Integer
        Dim CanalMostrarAlerta, CanalMostrarAlerta2 As String

        Mensaje = ""
        Correo = ""
        CorreoCopia = ""
        ErrorEnviarCorreo = False
        OrdenDeAmazon = ""
        CanalMostrarAlerta = ""

        Consulta = "select Valor from Parametro where CodigoParametro = 4" 'correo nivel 1
        CorreoAdministracion1 = mostrar.retornarcadena(Consulta, MyConString)

        Try
            CanalMostrarAlerta = ""
            Consulta = "select w.Url, isnull(Activo,0) as Activo from Alerta a, Webhook w " &
                    "where a.CodigoWebhook = w.CodigoWebhook " &
                    "and a.CodigoAlerta = 11"
            Tablita.Clear()
            mostrar.ejecuta_query_dt(Consulta, Tablita, MyConString)
            For Each dr As DataRow In Tablita.Rows
                Dim Activo As Boolean = CBool(dr("Activo").ToString)
                If Activo = True Then
                    If dr("Url").ToString <> "" Then
                        CanalMostrarAlerta = dr("Url").ToString
                    End If
                End If
            Next

            Consulta = "select OrdenDeAmazon from OrdenDeCompra where CorreoDeExencion is null"
            Tablita.Clear()
            mostrar.ejecuta_query_dt(Consulta, Tablita, MyConString)

            For Each row As DataRow In Tablita.Rows
                OrdenDeAmazon = row("OrdenDeAmazon").ToString

                If OrdenDeAmazon <> "" Then
                    ListaGuias = ""
                    ListaImpuestos = ""
                    TotalImpuesto = 0
                    ArchivosAdjuntos = ""
                    Mensaje = ""

                    'verifica si hay líneas de pedido en solicitado, tránsito o pendiente, si hay no genera guía ni envía el correo
                    Consulta = "select COUNT(1) from Pedido where OrdenDeAmazon = '" & OrdenDeAmazon & "' and CodigoEstadoPedido in (1,2,6) and Cantidad > 0"
                    NumSolicitadoTransito = mostrar.retornarentero(Consulta, MyConString)

                    If NumSolicitadoTransito = 0 Then 'aqui
                        'obtiene el impuesto a pagar de la órden de amazon
                        Consulta = "select isnull(SUM(Impuesto),0) as Impuesto from Pedido  where OrdenDeAmazon = '" & OrdenDeAmazon & "' and CodigoDeRastreo is not null "
                        TotalImpuesto = mostrar.retornardecimal(Consulta, MyConString)

                        Consulta = "select COUNT(1) from Pedido where OrdenDeAmazon = '" & OrdenDeAmazon & "' and Cantidad > 0 and CodigoEstadoPedido <> 4 "
                        NumSinCancelar = mostrar.retornarentero(Consulta, MyConString)

                        'si el correo tiene SiempreEnviarExencion = 1, aunque totalimpuesto = 0 se envía el correo
                        Consulta = "select COUNT(1) from CuentaDeCompra where Correo in ( " &
                                "select distinct Correo from Pedido where OrdenDeAmazon = '" & OrdenDeAmazon & "' and Correo is not null " &
                                ") and SiempreEnviarExencion = 1 "

                        'si el correo tiene NuncaEnviarExencion = 1, aunque totalimpuesto > 0 no se envía el correo
                        Consulta2 = "select COUNT(1) from CuentaDeCompra where Correo in ( " &
                                "select distinct Correo from Pedido where OrdenDeAmazon = '" & OrdenDeAmazon & "' and Correo is not null " &
                                ") and NuncaEnviarExencion = 1 "

                        If ((TotalImpuesto > 0 Or mostrar.retornarentero(Consulta, MyConString) > 0)) And mostrar.retornarentero(Consulta2, MyConString) = 0 And NumSinCancelar > 0 Then

                            Generar_Guias(OrdenDeAmazon)

                            If VerificarPaquete(OrdenDeAmazon, ListaGuias, ListaImpuestos) = True Then

                                Exito = True
                                texto = Split(ListaGuias, ",")
                                vectorImpuestos = Split(ListaImpuestos, ",")
                                For i = 0 To texto.Length - 1
                                    'verifica si hay alguna línea en el tracking que no haya sido cancelada y tengan cantidad, si todas las líneas del tracking están canceladas no genera el pdf de ese tracking
                                    Consulta = "select COUNT(1) from Pedido pe, Paquete pa where pe.CodigoPaquete = pa.CodigoPaquete and GuiaAerea = '" & texto(i) & "' and Cantidad > 0 and CodigoEstadoPedido <> 4 "
                                    If mostrar.retornarentero(Consulta, MyConString) > 0 Then 'genera el pdf
                                        If Exito = True Then
                                            NombreArchivo = ""
                                            Dim FormGuia As New GuiaPdf
                                            Exito = FormGuia.Crear_Reporte_Pdf(texto(i), NombreArchivo, MyConString, Mensaje)
                                            If NombreArchivo <> "" Then
                                                If ArchivosAdjuntos = "" Then
                                                    ArchivosAdjuntos = NombreArchivo
                                                Else
                                                    If InStr(ArchivosAdjuntos, NombreArchivo) = 0 Then
                                                        ArchivosAdjuntos = ArchivosAdjuntos + "," + NombreArchivo
                                                    End If

                                                End If
                                            End If

                                        End If
                                    End If
                                Next

                                If Exito = True Then

                                    Consulta = "select COUNT(*) from Pedido where OrdenDeAmazon = '" & OrdenDeAmazon & "' and Correo is not null"
                                    If mostrar.retornarentero(Consulta, MyConString) > 0 Then

                                        Correo = ""
                                        Consulta = "select top(1) isnull(pe.Correo,'') as Correo, isnull(cc.EnviarCopiaExencion,0) as EnviarCopiaExencion from Pedido pe, CuentaDeCompra cc where pe.Correo = cc.Correo and OrdenDeAmazon = '" & OrdenDeAmazon & "' and pe.Correo is not null"
                                        Tablita.Clear()
                                        mostrar.ejecuta_query_dt(Consulta, Tablita, MyConString)
                                        For Each dr As DataRow In Tablita.Rows
                                            Correo = dr("Correo").ToString
                                            EnviarCopiaExencion = CBool(dr("EnviarCopiaExencion").ToString)
                                            If EnviarCopiaExencion = True Then
                                                CorreoCopia = CorreoAdministracion1
                                            Else
                                                CorreoCopia = ""
                                            End If

                                        Next


                                        'enviar correo
                                        Subject = "Tax exempt for " & Correo

                                        Contenido = "Good day, attached proof of export.<br/><br/>" &
                                            "E-mail account: " & Correo & "<br/>" &
                                            "Order numbers:<br/>" &
                                            OrdenDeAmazon & "<br/><br/><br/>" &
                                            "Best Regards,<br/>" &
                                            "Mario Porres<br/><br/>"


                                        If ArchivosAdjuntos <> "" Then
                                            If enviar.Enviar_Correo_Con_Attachment_Verificar_Error(Correo, "tax-exempt@amazon.com", Subject, Contenido, CorreoCopia, ArchivosAdjuntos) = True Then
                                                'actualizar campo CorreoDeExencion
                                                Consulta = "Update OrdenDeCompra set CorreoDeExencion = GETDATE() Where OrdenDeAmazon = '" & OrdenDeAmazon & "'"
                                                mostrar.insertarmodificareliminar(Consulta, MyConString)
                                            Else

                                                Subject = "Error al enviar el correo en el procedimiento Correo Exencion "
                                                Contenido = "Error procedimiento Correo Exencion al enviar el correo desde: " + Correo + ", Orden de Amazon: " + OrdenDeAmazon + "<br/><br/>" + Contenido

                                                If CanalMostrarAlerta <> "" Then
                                                    EnviarAlertaSlackPromocion("ALERTA-EXENCION", Contenido, Subject, "Generar_CorreoExencion", CanalMostrarAlerta)
                                                End If
                                                'enviar.Enviar_Correo("cpx@guatemaladigital.net", CorreoAdministracion1, Subject, Contenido, "")
                                                'enviar.Enviar_Correo("cpx@guatemaladigital.net", "archivoslab@yahoo.es", Subject, Contenido, "")
                                            End If
                                        End If
                                    End If
                                Else 'Tarea 245
                                    Subject = "Error procedimiento Correo de Exencion"
                                    Contenido = "Orden Amazon: " + OrdenDeAmazon + ", " + Mensaje
                                    'enviar.Enviar_Correo("cpx@guatemaladigital.net", "archivoslab@yahoo.es", Subject, Contenido, "")
                                    If CanalMostrarAlerta <> "" Then
                                        EnviarAlertaSlackPromocion("ALERTA-EXENCION", Contenido, Subject, "Generar_CorreoExencion", CanalMostrarAlerta)
                                    End If

                                End If

                            End If ' if verificar paquete

                        Else
                            'como la suma de impuestos = 0, no se envía correo solo se actualiza campo CorreoDeExencion
                            'actualizar campo CorreoDeExencion 
                            'Consulta = "Update OrdenDeCompra set CorreoDeExencion = GETDATE() Where OrdenDeAmazon = '" & OrdenDeAmazon & "'"
                            'mostrar.insertarmodificareliminar(Consulta, MyConString)

                        End If 'if totalimpuesto <> 0

                    End If 'NumSolicitadoTransito

                    'End If 'if verifica si hay lineas con Shipment planned o Shipping soon, si hay no envía correo
                End If 'if orden de amazon

            Next

        Catch ex As Exception

            Consulta = "select url from Webhook where CodigoWebhook = 1" 'canal de alerta sistemas
            CanalMostrarAlerta = mostrar.retornarcadena(Consulta, MyConString)

            Subject = "Error procedimiento Correo de Exencion"
            Contenido = "Orden Amazon: " + OrdenDeAmazon + Chr(13) + ex.Message.ToString
            'enviar.Enviar_Correo("cpx@guatemaladigital.net", "archivoslab@yahoo.es", Subject, Contenido, "")
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
        'jSonSlack += "            ""title"": """ + " " + """, "
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
            'webClient.UploadString("https://hooks.slack.com/services/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", json)
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12
            webClient.UploadString(UrlCanalSlack, json)
        Catch ex As Exception
            Cadena = "Error al enviar mensaje Slack :  " + "<br/><br/>" + ex.Message.ToString + "<br/><br/>"
            'Guardar_Datos_Archivo_Texto(Cadena)
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
        Consulta = "select COUNT(1) from Pedido where OrdenDeAmazon = '" & OrdenDeAmazon & "' and CodigoEstadoPedido = 3 and Cantidad > 0"

        If mostrar.retornarentero(Consulta, MyConString) > 0 Then 'la orden tienen algún producto recibido
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
        Dim TrackingNulo As Integer
        Dim GuiaAerea As String

        'número de líneas de la orden que no tienen tracking (codigoderastreo = null)
        Consulta = "select COUNT(1) from Pedido where OrdenDeAmazon = '" & OrdenDeAmazon & "' and (CodigoDeRastreo is null or LEN(codigoderastreo) = 0)"
        TrackingNulo = mostrar.retornarentero(Consulta, MyConString)

        'total de tracking de la orden
        Consulta = "select COUNT(1) from ( " &
            "select distinct CodigoDeRastreo from Pedido where OrdenDeAmazon = '" & OrdenDeAmazon & "' and CodigoDeRastreo is not null and CodigoEstadoPedido <> 4 " &
            ") z "

        TotalPedido = mostrar.retornarentero(Consulta, MyConString)
        TotalPaquete = 0
        Exito = False
        ListaGuias = ""
        ListaImpuestos = ""

        If TotalPedido > 0 Then 'And TrackingNulo = 0

            'Consulta = "select distinct CodigoDeRastreo, case when CHARINDEX('(', CodigoDeRastreo) > 0 and CHARINDEX(')', CodigoDeRastreo) > 0 then SUBSTRING(CodigoDeRastreo,CHARINDEX('(', CodigoDeRastreo) + 1, CHARINDEX(')', CodigoDeRastreo) - CHARINDEX('(', CodigoDeRastreo) - 1) else CodigoDeRastreo end as Tracking from Pedido where OrdenDeAmazon = '" & OrdenDeAmazon & "'  "
            Consulta = "select CodigoDeRastreo, case when CHARINDEX('(', CodigoDeRastreo) > 0 and CHARINDEX(')', CodigoDeRastreo) > 0 then SUBSTRING(CodigoDeRastreo,CHARINDEX('(', CodigoDeRastreo) + 1, CHARINDEX(')', CodigoDeRastreo) - CHARINDEX('(', CodigoDeRastreo) - 1) else CodigoDeRastreo end as Tracking, isnull(SUM(Impuesto),0) as Impuesto from Pedido  where OrdenDeAmazon = '" & OrdenDeAmazon & "' and CodigoEstadoPedido <> 4 group by CodigoDeRastreo having CodigoDeRastreo is not null "

            Using mySqlConnection2 As New System.Data.SqlClient.SqlConnection(MyConString)
                mySqlConnection2.Open()
                Dim mySqlCommand2 As New System.Data.SqlClient.SqlCommand(Consulta, mySqlConnection2)
                Dim myDataReader2 As Data.SqlClient.SqlDataReader
                myDataReader2 = mySqlCommand2.ExecuteReader()

                Do While myDataReader2.Read()
                    If myDataReader2.IsDBNull(1) = False Then
                        CodigoDeRastreo = myDataReader2.GetString(1)

                        Consulta = "select COUNT(1) from Paquete where codigoderastreo like '%" & CodigoDeRastreo & "%' "
                        If mostrar.retornarentero(Consulta, MyConString) = 1 Then
                            TotalPaquete = TotalPaquete + 1
                            Consulta = "select isnull(GuiaAerea,'') from Paquete where codigoderastreo like '%" & CodigoDeRastreo & "%' "
                            GuiaAerea = mostrar.retornarcadena(Consulta, MyConString)
                            GuiaAerea = Trim(GuiaAerea)

                            If GuiaAerea = "" Then
                                'GuiaAerea = Calcular_Guia_Aerea(CodigoDeRastreo)
                                GuiaAerea = Obtener_Nueva_Guia()
                                Consulta = "update paquete set GuiaAerea = '" & GuiaAerea & "', Generado = 1 where codigoderastreo like '%" & CodigoDeRastreo & "%' "
                                mostrar.insertarmodificareliminar(Consulta, MyConString)
                            End If

                            If ListaGuias = "" Then
                                ListaGuias = GuiaAerea
                            Else
                                ListaGuias = ListaGuias & "," & GuiaAerea
                            End If
                        End If
                    End If

                    If myDataReader2.IsDBNull(2) = False Then
                        If ListaImpuestos = "" Then
                            ListaImpuestos = CStr(myDataReader2.GetDecimal(2))
                        Else
                            ListaImpuestos = ListaImpuestos & "," & CStr(myDataReader2.GetDecimal(2))
                        End If
                    End If

                Loop

                'Consulta = "select isnull(SUM(t.Impuesto),0) from ( " & _
                '                "select CodigoDeRastreo, case when CHARINDEX('(', CodigoDeRastreo) > 0 and CHARINDEX(')', CodigoDeRastreo) > 0 then SUBSTRING(CodigoDeRastreo,CHARINDEX('(', CodigoDeRastreo) + 1, CHARINDEX(')', CodigoDeRastreo) - CHARINDEX('(', CodigoDeRastreo) - 1) else CodigoDeRastreo end as Tracking, isnull(SUM(Impuesto),0) as Impuesto from Pedido  where OrdenDeAmazon = '" & OrdenDeAmazon & "' group by CodigoDeRastreo having CodigoDeRastreo is not null " & _
                '            ") t"

                'TotalImpuesto = mostrar.retornardecimal(Consulta, MyConString)

                If TotalPedido = TotalPaquete Then
                    Exito = True
                End If

                myDataReader2.Close()
                mySqlConnection2.Close()
            End Using

        End If

        VerificarPaquete = Exito
    End Function

    'Function Calcular_Guia_Aerea(ByVal CodigoDeRastreo As String) As String

    '    Dim CadFechaOrden, Guia As String
    '    Dim Contador As Integer
    '    Dim Exito As Boolean

    '    'obtiene los datos que se guardarán en la tabla paquete
    '    Consulta = "select top(1) CONVERT(date,FechaDeOrden)  " &
    '                    "from Pedido where CodigoDeRastreo like '%" & CodigoDeRastreo & "%'"
    '    CadFechaOrden = mostrar.retornafechaSinHora(Consulta, MyConString)
    '    CadFechaOrden = CadFechaOrden.Substring(6, 4) & "-" & CadFechaOrden.Substring(3, 2) & "-" & CadFechaOrden.Substring(0, 2)



    '    Consulta = "select top(1) GuiaAerea from Paquete where  Fecha <= '" & CadFechaOrden & "' and GuiaAerea is not null order by Fecha desc, GuiaAerea desc"
    '    Guia = mostrar.retornarcadena(Consulta, MyConString)


    '    Contador = 0
    '    Exito = False
    '    While Not Exito
    '        Guia = CStr(CInt(Guia) - 1)
    '        Consulta = "Select count(*) from Paquete Where GuiaAerea = '" & Guia & "'"
    '        If mostrar.retornarentero(Consulta, MyConString) = 0 Then
    '            Exito = True
    '        Else
    '            Contador = Contador + 1
    '        End If
    '    End While

    '    Calcular_Guia_Aerea = Guia
    'End Function


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
        Consulta = "select distinct pe.CodigoDeRastreo, (select COUNT(*) from Pedido where CodigoEstadoPedido = 3 and CodigoDeRastreo = pe.CodigoDeRastreo) as Recibidos, " &
             "case when CHARINDEX('(', pe.CodigoDeRastreo) > 0 and CHARINDEX(')', pe.CodigoDeRastreo) > 0 then SUBSTRING(pe.CodigoDeRastreo,CHARINDEX('(', pe.CodigoDeRastreo) + 1, CHARINDEX(')', pe.CodigoDeRastreo) - CHARINDEX('(', pe.CodigoDeRastreo) - 1) else pe.CodigoDeRastreo end as TrackingCorto " &
             "from Pedido pe where pe.OrdenDeAmazon = '" & OrdenDeAmazon & "' and CodigoDeRastreo is not null and CodigoEstadoPedido <> 4 "
        Tablita.Clear()
        mostrar.ejecuta_query_dt(Consulta, Tablita, MyConString)


        For Each row As DataRow In Tablita.Rows
            CodigoDeRastreo = row(2).ToString

            'verifica que el tracking no exista en la tabla paquete
            Consulta = "Select Count(*) from Paquete where CodigoDeRastreo like '%" & CodigoDeRastreo & "%'"
            If mostrar.retornarentero(Consulta, MyConString) <> 0 Then
                'verifica si existe registros en la tabla pedido con el código de rastreo y sin número de paquete (sin guía)

                Consulta = "select COUNT(*) from Pedido where CodigoDeRastreo like '%" & CodigoDeRastreo & "%' and CodigoPaquete is null"
                If mostrar.retornarentero(Consulta, MyConString) = 0 Then Return

                'agrega el código de paquete al registro de pedido

                Consulta = "Select CodigoPaquete from Paquete where CodigoDeRastreo like '%" & CodigoDeRastreo & "%'"
                    CodigoPaquete = mostrar.retornarentero(Consulta, MyConString)

                    Consulta = "update Pedido set CodigoPaquete = " & CStr(CodigoPaquete) & " where CodigoDeRastreo in ( " &
                                        "select top(1) CodigoDeRastreo from Pedido where CodigoDeRastreo = '" & CodigoDeRastreo & "' or CodigoDeRastreo like '%" & CodigoDeRastreo & "%' " &
                                    ")"

                mostrar.insertarmodificareliminar(Consulta, MyConString)

                Return
            End If

            'verifica si el tracking existe en tabla pedido
            Consulta = "select Count(*) from Pedido where CodigoDeRastreo like '%" & CodigoDeRastreo & "%'"
            If mostrar.retornarentero(Consulta, MyConString) = 0 Then Return

            'obtiene los datos que se guardarán en la tabla paquete
            Consulta = "select top(1) CONVERT(date,FechaDeOrden)  " &
                        "from Pedido where CodigoDeRastreo like '%" & CodigoDeRastreo & "%'"
            CadFechaOrden = mostrar.retornafechaSinHora(Consulta, MyConString)
            CadFechaOrden = CadFechaOrden.Substring(6, 4) & "-" & CadFechaOrden.Substring(3, 2) & "-" & CadFechaOrden.Substring(0, 2)

            Consulta = "DECLARE @valores VARCHAR(1000) " &
                            "select @valores= COALESCE(@valores + ', ', '') + t.TipoDeProducto from ( " &
                            "select distinct(isnull(TipoDeProducto,'')) as TipoDeProducto from Pedido where CodigoDeRastreo like '%" & CodigoDeRastreo & "%' " &
                            ") t " &
                            "select replace(@valores,',',' ') "

            Descripcion = mostrar.retornarcadena(Consulta, MyConString)

            Consulta = "select top(1) isnull(Vendedor,'') from Pedido where CodigoDeRastreo like '%" & CodigoDeRastreo & "%'"
            EnviadoPor = mostrar.retornarcadena(Consulta, MyConString)

            'verifica si hay productos sin peso
            'si hay algún producto sin peso (NULL), Peso = 0
            'si todos los productos tienen peso, se suman todos
            Consulta = "select COUNT(*) from Producto where CodigoAmazon in ( " &
                                    "select CodigoAmazon from Pedido where CodigoDeRastreo like '%" & CodigoDeRastreo & "%' " &
                                    ") and Peso is null "

            If mostrar.retornarentero(Consulta, MyConString) = 0 Then
                Consulta = "select sum(isnull(Peso,0)) from Producto where CodigoAmazon in ( " &
                                    "select CodigoAmazon from Pedido where CodigoDeRastreo like '%" & CodigoDeRastreo & "%' " &
                                    ")"
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
            Consulta = "select top(1) pr.CodigoEstablecimiento " &
                        "from Pedido pe, Producto pr where pe.CodigoAmazon = pr.CodigoAmazon and pe.OrdenDeAmazon = '" & OrdenDeAmazon & "' and CodigoDeRastreo is not null "
            Tablita.Clear()
            mostrar.ejecuta_query_dt(Consulta, Tablita, MyConString)

            CodigoEstablecimiento = Tablita.Rows(0)("CodigoEstablecimiento").ToString

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
            Consulta = "Select count(*) from Paquete Where GuiaAerea = '" & "G" & Guia & "'"
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


        Return "G" & Guia
    End Function


    Sub Insertar_Paquete(ByVal Tracking As String, ByVal GuiaAerea As String, ByVal EnviadoPor As String, ByVal Descripcion As String, ByVal Peso As String, ByVal PesoVolumetrico As String, ByVal MontoImpuesto As String, ByVal Generado As String, ByVal EstadoPaquete As String, ByVal CodigoEstablecimiento As String, ByRef CodigoPaquete As Integer)


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


        Consulta = "select Count(*) from Paquete"
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

    Sub Actualizar_Estado_Paquete_Venta(ByVal CodigoPaquete As Integer, ByVal Tracking As String, ByVal WareHouseNotificacion As Boolean)

        Dim TrackingReducido, CadTracking, Mensaje, MensajeError As String
        Dim dtventas As New DataTable

        TrackingReducido = "" : Mensaje = "" : MensajeError = ""
        Validar_Tracking(Tracking, TrackingReducido, MensajeError)

        If TrackingReducido <> "" Then
            CadTracking = TrackingReducido
        Else
            CadTracking = Tracking
        End If


        If MensajeError <> "Se encontró más de 1 tracking" Then
            'verifica que no exista CodigoPaquete en la tabla Pedido
            'Consulta = "Select top(1) CodigoPedido from Pedido Where CodigoPaquete = " & CStr(CodigoPaquete)
            'If mostrar.retornavacio(Consulta, MyConString) = 0 Then

            Consulta = "update Pedido set CodigoPaquete = " & CStr(CodigoPaquete) & " where CodigoDeRastreo in ( " &
                            "select top(1) CodigoDeRastreo from Pedido where CodigoDeRastreo = '" & CadTracking & "' or CodigoDeRastreo like '%" & CadTracking & "%' " &
                        ")"

            mostrar.insertarmodificareliminar(Consulta, MyConString)

            'Actualizacion y envio de correo de estado de entrega
            Consulta = "select distinct vp.CodigoVenta from VentaPedido vp inner join Pedido p on vp.CodigoPedido = p.CodigoPedido inner join Venta v on vp.CodigoVenta = v.CodigoVenta "
            Consulta += "inner join Paquete pa on p.CodigoPaquete = pa.CodigoPaquete where v.CodigoEstadoEntrega = 2 and pa.GuiaAerea is not null and p.CodigoPaquete = " + CodigoPaquete.ToString.Trim
            mostrar.ejecuta_query_dt(Consulta, dtventas, MyConString)

            If dtventas.Rows.Count > 0 Then
                Try
                    For Each fila As DataRow In dtventas.Rows
                        'Actualiza el codigoestadoentrega 3 en venta
                        Consulta = "update Venta set CodigoEstadoEntrega = 3 where CodigoVenta = " + fila("CodigoVenta").ToString + " and CodigoEstadoEntrega = 2"
                        mostrar.insertarmodificareliminar(Consulta, MyConString)

                        'Envio de correo de estado de entrega
                        enviar.Enviar_Correo_Rastreo(fila("CodigoVenta").ToString, MyConString)
                    Next
                Catch ex As Exception

                End Try
            End If
            'Fin Actualizacion

            If WareHouseNotificacion = True Then
                Consulta = "update Pedido set CodigoEstadoPedido = 2 where CodigoDeRastreo in ( " &
                                "select top(1) CodigoDeRastreo from Pedido where CodigoDeRastreo = '" & CadTracking & "' or CodigoDeRastreo like '%" & CadTracking & "%' " &
                            ") and CodigoEstadoPedido = 1 "

                mostrar.insertarmodificareliminar(Consulta, MyConString)

            End If

            'End If 'No se encontraron registros con codigopaquete
        End If 'mensajeerror

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
            If mostrar.retornarentero(Consulta, MyConString) = 0 Then 'no hay tracking
                'verifica si el tracking es númerico (puede ser de Fedex)
                If IsNumeric(Tracking) = False Then
                    MensajeError = "El tracking no existe"
                    Exito = False
                Else 'toma los últimos 12 dígitos del tracking (por si es de Fedex)
                    TrackingReducido = Microsoft.VisualBasic.Right(Tracking, 12)
                    'verifica si existe el tracking 
                    Consulta = "select Count(*)  from Pedido where (CodigoDeRastreo = '" & TrackingReducido & "' or CodigoDeRastreo like '%" & TrackingReducido & "%')  and CodigoEstadoPedido <> 4"
                    If mostrar.retornarentero(Consulta, MyConString) = 0 Then 'no hay tracking
                        MensajeError = "El tracking no existe"
                        Exito = False
                    Else
                        'verifica si se encontró más de 1 tracking
                        'Consulta = "select COUNT(*) from (select distinct(CodigoDeRastreo) from Pedido where (CodigoDeRastreo = '" & TrackingReducido & "' or CodigoDeRastreo like '%" & TrackingReducido & "%') and CodigoEstadoPedido <> 4 ) t"
                        Consulta = "select COUNT(1) from (select distinct replace(SUBSTRING(codigoderastreo, (CHARINDEX('(', codigoderastreo) + 1), (LEN(codigoderastreo) - CHARINDEX('(', codigoderastreo))),')','') as CodigoDeRastreo "
                        Consulta += "from Pedido where (CodigoDeRastreo = '" & TrackingReducido & "' or CodigoDeRastreo like '%" & TrackingReducido & "%') and CodigoEstadoPedido <> 4 ) t"
                        If mostrar.retornarentero(Consulta, MyConString) > 1 Then
                            MensajeError = "Se encontró más de 1 tracking"
                            Exito = False
                        End If

                    End If
                End If

            Else
                'verifica si se encontró más de 1 tracking
                'Consulta = "select COUNT(*) from (select distinct(CodigoDeRastreo) from Pedido where (CodigoDeRastreo = '" & Tracking & "' or CodigoDeRastreo like '%" & Tracking & "%') and CodigoEstadoPedido <> 4 ) t"
                Consulta = "select COUNT(1) from (select distinct replace(SUBSTRING(codigoderastreo, (CHARINDEX('(', codigoderastreo) + 1), (LEN(codigoderastreo) - CHARINDEX('(', codigoderastreo))),')','') as CodigoDeRastreo "
                Consulta += "from Pedido where (CodigoDeRastreo = '" & Tracking & "' or CodigoDeRastreo like '%" & Tracking & "%') and CodigoEstadoPedido <> 4 ) t"
                If mostrar.retornarentero(Consulta, MyConString) > 1 Then
                    MensajeError = "Se encontró más de 1 tracking"
                    Exito = False
                End If
            End If 'no hay tracking

        End If 'longitud > 40

        Validar_Tracking = Exito
    End Function

End Module

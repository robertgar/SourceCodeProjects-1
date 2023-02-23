using System.Data;
using System.Text;
using System.Security.Cryptography;

namespace Principal {
    internal class Maker {
        internal class Vars {
            public connection.Slack slack = new connection.Slack();
            public connection.SQL execute = new connection.SQL();
            public common.UseCommon use = new common.UseCommon();
        }
        bool isSimulation;
        public void MakeAll(ref bool isSimulation) {
            this.isSimulation = isSimulation;
            Vars js = new Vars();
            DateTime _now = DateTime.Now;
            int[] assessedCharges = { 0, 0, 0, 0, 0 };

            js.slack.data.Clear();
            js.slack.data.alertTitle.Append("Alert");
            js.slack.data.subject.Append("Procedure alert info: the process has begun...");
            js.slack.data.message.Append("Start time: ").Append(_now);
            js.slack.data.warningColour = js.slack.data.getColour.Green;
            js.slack.send();

            try {
                TryMake(ref js, ref assessedCharges);
            } catch (Exception e) {
                js.slack.sendError(e.ToString());
            }

            js.slack.data.Clear();
            js.slack.data.alertTitle.Append("Alert");
            js.slack.data.subject.Append("Procedure alert info: the process has been completed!");
            js.slack.data.message.Append("Outstanding receivables: ").AppendLine(assessedCharges[0].ToString());
            js.slack.data.message.Append("Collections made: ").AppendLine(assessedCharges[1].ToString());
            js.slack.data.message.Append("Expired cards: ").AppendLine(assessedCharges[2].ToString());
            js.slack.data.message.Append("Error collections: ").AppendLine(assessedCharges[3].ToString());
            js.slack.data.message.Append("Total records: ").AppendLine(assessedCharges[4].ToString());
            js.slack.data.message.Append("Total execution time: ").Append(DateTime.Now - _now);
            js.slack.data.warningColour = js.slack.data.getColour.Green;
            js.slack.send();
        }

        /// <summary>
        /// Main method of collection
        /// </summary>
        /// <param name="js"></param>
        /// <param name="assessedCharges"></param>
        private void TryMake(ref Vars js, ref int[] assessedCharges) {
            js.use.Clear();
            js.use.query.AppendLine("declare @Parameter int = (Select Valor from Parametro where CodigoParametro = 17)");
            js.use.query.AppendLine("select");
            js.use.query.AppendLine("    cl.CodigoCliente as CustomerCode,");
            js.use.query.AppendLine("    case");
            js.use.query.AppendLine("        when datediff(");
            js.use.query.AppendLine("            month,");
            js.use.query.AppendLine("            getdate(),");
            js.use.query.AppendLine("            concat(");
            js.use.query.AppendLine("                cl.VencimientoTarjeta % 10000,");
            js.use.query.AppendLine("                '-',");
            js.use.query.AppendLine("                (cl.VencimientoTarjeta - cl.VencimientoTarjeta % 10000) / 10000,");
            js.use.query.AppendLine("                '-01'");
            js.use.query.AppendLine("            )");
            js.use.query.AppendLine("        ) < 0 then -1");
            js.use.query.AppendLine("        when datediff(");
            js.use.query.AppendLine("            month,");
            js.use.query.AppendLine("            getdate(),");
            js.use.query.AppendLine("            concat(");
            js.use.query.AppendLine("                cl.VencimientoTarjeta % 10000,");
            js.use.query.AppendLine("                '-',");
            js.use.query.AppendLine("                (cl.VencimientoTarjeta - cl.VencimientoTarjeta % 10000) / 10000,");
            js.use.query.AppendLine("                '-01'");
            js.use.query.AppendLine("            )");
            js.use.query.AppendLine("        ) < 2 then 0");
            js.use.query.AppendLine("        when datediff(");
            js.use.query.AppendLine("            day,");
            js.use.query.AppendLine("            getdate(),");
            js.use.query.AppendLine("            isnull(");
            js.use.query.AppendLine("                max(t.Fecha),");
            js.use.query.AppendLine("                dateadd(day, -@Parameter, getdate())");
            js.use.query.AppendLine("            )");
            js.use.query.AppendLine("            ) >= @Parameter then 0");
            js.use.query.AppendLine("        else 1");
            js.use.query.AppendLine("    end as Charging,");
            js.use.query.AppendLine("    cl.NumeroTarjeta as CardNumber,");
            js.use.query.AppendLine("    concat(");
            js.use.query.AppendLine("        cl.VencimientoTarjeta % 10000,");
            js.use.query.AppendLine("        '-',");
            js.use.query.AppendLine("        (cl.VencimientoTarjeta - cl.VencimientoTarjeta % 10000) / 10000,");
            js.use.query.AppendLine("        '-01'");
            js.use.query.AppendLine("    ) as ExpireDate,");
            js.use.query.AppendLine("    cl.Correo as MailCustomer,");
            js.use.query.AppendLine("    convert(decimal(7, 4), sum(c.CPC)) as Total");
            js.use.query.AppendLine("from");
            js.use.query.AppendLine("    Click as c");
            js.use.query.AppendLine("    inner join Anuncio as a on a.CodigoAnuncio = c.CodigoAnuncio");
            js.use.query.AppendLine("    inner join Cliente as cl on cl.CodigoCliente = a.CodigoCliente");
            js.use.query.AppendLine("    left join(");
            js.use.query.AppendLine("        select");
            js.use.query.AppendLine("            v.CodigoCliente,");
            js.use.query.AppendLine("            max(v.Fecha) as Fecha");
            js.use.query.AppendLine("        from");
            js.use.query.AppendLine("            Venta as v");
            js.use.query.AppendLine("            inner join DetalleVentaServicio as dvs");
            js.use.query.AppendLine("                on dvs.CodigoVenta = v.CodigoVenta");
            js.use.query.AppendLine("        where");
            js.use.query.AppendLine("            v.CodigoProducto = 0");
            js.use.query.AppendLine("            and dvs.CodigoServicio = 5");
            js.use.query.AppendLine("        group by");
            js.use.query.AppendLine("            v.CodigoCliente");
            js.use.query.AppendLine("    ) as t on t.CodigoCliente = cl.CodigoCliente");
            js.use.query.AppendLine("where");
            js.use.query.AppendLine("    a.CodigoEstadoDeAnuncio = 1");
            js.use.query.AppendLine("    and isnull(a.Aprobado, 0) = 1");
            js.use.query.AppendLine("    and c.CodigoVenta is null");
            js.use.query.AppendLine("    and isnull(cl.AnunciosSuspendidos, 0) = 0");
            js.use.query.AppendLine("group by");
            js.use.query.AppendLine("    cl.CodigoCliente,");
            js.use.query.AppendLine("    cl.VencimientoTarjeta,");
            js.use.query.AppendLine("    cl.NumeroTarjeta,");
            js.use.query.AppendLine("    cl.Correo");

            DataTable tabPrincipal = new DataTable();
            js.execute.fillTable(ref js.use.query, ref tabPrincipal);
            assessedCharges[4] = tabPrincipal.Rows.Count;

            if (tabPrincipal.Rows.Count == 0) { return; }

            foreach (DataRow row in tabPrincipal.Rows) {
                switch (int.Parse(row["Charging"].ToString())) {
                    case 1:
                        //Next turn
                        assessedCharges[0]++;
                        break;
                    case 0:
                        //Create invoice
                        //Make collect
                        //Facturar y enviar factura
                        ChargingAdd(row, ref js, ref assessedCharges);
                        break;
                    case -1:
                        assessedCharges[2]++;
                        //Tarjeta vencida
                        //Enviar correo y solicitar renovación de tarjeta (Ver gd-plus)
                        string result = sentMailCardExpire(row["MailCustomer"].ToString(), row["CardNumber"].ToString(), ref js);
                        if (!result.Contains("Error:")) { continue; }

                        js.slack.sendError(result);
                        break;
                }
            }
        }

        private void ChargingAdd(DataRow row, ref Vars js, ref int[] assessedCharges) {
            string Invoice = getNewInvoice(row["CustomerCode"].ToString(), ref js);
            string _cardNumer = getNumerotarjeta(row["CardNumber"].ToString());
            string _year = DateTime.Parse(row["ExpireDate"].ToString()).Year.ToString();
            string _month = DateTime.Parse(row["ExpireDate"].ToString()).Month.ToString();
            string _total = row["Total"].ToString();

            EnvioData _data = new EnvioData(
                _total,
                _cardNumer,
                _month,
                _year,
                "",
                Invoice
            );
            PagoTarjeta _payment = CobroPeti.AutoCredomatic(_data);
            string result;

            if (_payment.Code == 100) {
                var service = new wsGD.ServiceSoapClient();
                result = service.fuFacturar(Invoice.ToString(), "", true);

                assessedCharges[result.Contains("Error") ? 3 : 1]++;
            } else {
                result = _payment.Response.respuestatext;
                assessedCharges[1]++;
            }

            Guardar_Datos_Archivo_Texto("CodigoFactura: " + Invoice + " Mensaje:" + result);
        }

        /// <summary>
        /// Creates invoice by customer code
        /// creates sales by advertisement and creates a bill for the total of sales
        /// </summary>
        /// <param name="CustomerCode"/>
        /// <param name="js"/>
        /// <returns></returns>
        private string getNewInvoice(string CustomerCode, ref Vars js) {
            js.use.Clear();
            js.use.query.AppendLine("begin tran begin try");
            //Create variables
            js.use.query.Append("declare @CustomerCode int = ").AppendLine(CustomerCode);
            js.use.query.AppendLine("declare @InvoiceCode int = (select Max(CodigoFactura) + 1 from Factura)");
            js.use.query.AppendLine("declare @customer_data table(");
            js.use.query.AppendLine("    _code int,");
            js.use.query.AppendLine("    _name varchar(200),");
            js.use.query.AppendLine("    _mail varchar(255),");
            js.use.query.AppendLine("    _address varchar(500),");
            js.use.query.AppendLine("    _phone varchar(50),");
            js.use.query.AppendLine("    _nit  varchar(25)");
            js.use.query.AppendLine(")");
            js.use.query.AppendLine("declare @card_data table(");
            js.use.query.AppendLine("    _number varchar(256),");
            js.use.query.AppendLine("    _expiry int,");
            js.use.query.AppendLine("    _name varchar(200),");
            js.use.query.AppendLine("    _code varbinary(256),");
            js.use.query.AppendLine("    _type varchar(50)");
            js.use.query.AppendLine(")");
            js.use.query.AppendLine("declare @advert_data table(");
            js.use.query.AppendLine("    _name varchar(50),");
            js.use.query.AppendLine("    _clicks int,");
            js.use.query.AppendLine("    _total decimal(5, 2)");
            js.use.query.AppendLine(")");

            //Fill variables
            js.use.query.AppendLine("insert into @customer_data");
            js.use.query.AppendLine("    select");
            js.use.query.AppendLine("        CodigoCliente,");
            js.use.query.AppendLine("        NombreFactura,");
            js.use.query.AppendLine("        Correo,");
            js.use.query.AppendLine("        DireccionFactura,");
            js.use.query.AppendLine("        Telefono,");
            js.use.query.AppendLine("        Nit");
            js.use.query.AppendLine("    from");
            js.use.query.AppendLine("        Cliente");
            js.use.query.AppendLine("    where");
            js.use.query.AppendLine("        CodigoCliente = @CustomerCode");
            js.use.query.AppendLine("insert into @card_data");
            js.use.query.AppendLine("    select");
            js.use.query.AppendLine("        NumeroTarjeta,");
            js.use.query.AppendLine("        VencimientoTarjeta,");
            js.use.query.AppendLine("        NombreTarjeta,");
            js.use.query.AppendLine("        Clave,");
            js.use.query.AppendLine("        TipoTarjeta");
            js.use.query.AppendLine("    from");
            js.use.query.AppendLine("        Cliente");
            js.use.query.AppendLine("    where");
            js.use.query.AppendLine("        CodigoCliente = @CustomerCode");
            js.use.query.AppendLine("insert into @advert_data");
            js.use.query.AppendLine("    select");
            js.use.query.AppendLine("        a.Nombre,");
            js.use.query.AppendLine("        count(cl.CodigoAnuncio),");
            js.use.query.AppendLine("        convert(decimal(7, 2), sum(cl.CPC))");
            js.use.query.AppendLine("    from");
            js.use.query.AppendLine("        Click as cl");
            js.use.query.AppendLine("        inner join Anuncio as a on a.CodigoAnuncio = cl.CodigoAnuncio");
            js.use.query.AppendLine("        inner join Cliente as c on c.CodigoCliente = a.CodigoCliente");
            js.use.query.AppendLine("    where");
            js.use.query.AppendLine("        isnull(c.AnunciosSuspendidos, 0) = 0");
            js.use.query.AppendLine("        and cl.CodigoVenta is null");
            js.use.query.AppendLine("        and isnull(a.Aprobado, 0) = 1");
            js.use.query.AppendLine("        and a.CodigoEstadoDeAnuncio = 1");
            js.use.query.AppendLine("        and c.CodigoCliente = @CustomerCode");
            js.use.query.AppendLine("    group by");
            js.use.query.AppendLine("        a.Nombre");

            //Insert invoice
            js.use.query.AppendLine("insert into Factura(");
            js.use.query.AppendLine("    CodigoFactura,");
            js.use.query.AppendLine("    SerieDeFactura,");
            js.use.query.AppendLine("    NumeroDeFactura,");
            js.use.query.AppendLine("    Nit,");
            js.use.query.AppendLine("    Nombre,");
            js.use.query.AppendLine("    Direccion,");
            js.use.query.AppendLine("    CodigoEmpresa,");
            js.use.query.AppendLine("    CodigoEstadoFactura,");
            js.use.query.AppendLine("    Fecha,");
            js.use.query.AppendLine("    Total,");
            js.use.query.AppendLine("    TotalSinComisiones,");
            js.use.query.AppendLine("    Observaciones,");
            js.use.query.AppendLine("    CodigoCliente,");
            js.use.query.AppendLine("    TotalFactura,");
            js.use.query.AppendLine("    Telefonos,");
            js.use.query.AppendLine("    TotalPendiente,");
            js.use.query.AppendLine("    MontoServicioEnEfectivo,");
            js.use.query.AppendLine("    NombreRecibido,");
            js.use.query.AppendLine("    EsFacturaElectronica,");
            js.use.query.AppendLine("    DatosDeReembolsoPendientes,");
            js.use.query.AppendLine("    CodigoUsuarioDespacho,");
            js.use.query.AppendLine("    Envio,");
            js.use.query.AppendLine("    FechaPago");
            js.use.query.AppendLine(")values(");
            js.use.query.AppendLine("    @InvoiceCode,");
            js.use.query.AppendLine("    'Temp',");
            js.use.query.AppendLine("    1000000000,");
            js.use.query.AppendLine("    (select _nit from @customer_data),");
            js.use.query.AppendLine("    (select _name from @customer_data),");
            js.use.query.AppendLine("    (select _address from @customer_data),");
            js.use.query.AppendLine("    1,");
            js.use.query.AppendLine("    0,");
            js.use.query.AppendLine("    getdate(),");
            js.use.query.AppendLine("    (select sum(_total) from @advert_data),");
            js.use.query.AppendLine("    (select sum(_total) from @advert_data),");
            js.use.query.AppendLine("    '',");
            js.use.query.AppendLine("    (select _code from @customer_data),");
            js.use.query.AppendLine("    (select sum(_total) from @advert_data),");
            js.use.query.AppendLine("    (select _phone from @customer_data),");
            js.use.query.AppendLine("    0.00,");
            js.use.query.AppendLine("    0.00,");
            js.use.query.AppendLine("    'Auto',");
            js.use.query.AppendLine("    1,");
            js.use.query.AppendLine("    0,");
            js.use.query.AppendLine("    0,");
            js.use.query.AppendLine("    0.00,");
            js.use.query.AppendLine("    getdate()");
            js.use.query.AppendLine(")");

            //Insert sale
            js.use.query.AppendLine("insert into Venta(");
            js.use.query.AppendLine("    Fecha,");
            js.use.query.AppendLine("    NombreCliente,");
            js.use.query.AppendLine("    CorreoCliente,");
            js.use.query.AppendLine("    Cantidad,");
            js.use.query.AppendLine("    CodigoProducto,");
            js.use.query.AppendLine("    NombreProducto,");
            js.use.query.AppendLine("    Monto,");
            js.use.query.AppendLine("    CodigoFormaDePago,");
            js.use.query.AppendLine("    DireccionDeEntrega,");
            js.use.query.AppendLine("    Factura,");
            js.use.query.AppendLine("    Confirmada,");
            js.use.query.AppendLine("    Telefonos,");
            js.use.query.AppendLine("    NumeroTarjeta,");
            js.use.query.AppendLine("    VencimientoTarjeta,");
            js.use.query.AppendLine("    NombreTarjeta,");
            js.use.query.AppendLine("    Cuotas,");
            js.use.query.AppendLine("    MontoCuota,");
            js.use.query.AppendLine("    EnPreorden,");
            //js.use.query.AppendLine("    -- CodigoSeguridad,");
            js.use.query.AppendLine("    CodigoRedCrediticia,");
            js.use.query.AppendLine("    CodigoTipoDeTarjeta,");
            js.use.query.AppendLine("    Pendiente,");
            js.use.query.AppendLine("    CodigoEstadoDeVenta,");
            js.use.query.AppendLine("    ProductoVerificado,");
            js.use.query.AppendLine("    CodigoFactura,");
            js.use.query.AppendLine("    MontoProveedor,");
            js.use.query.AppendLine("    CodigoCliente,");
            js.use.query.AppendLine("    Despachada,");
            js.use.query.AppendLine("    TotalVenta,");
            js.use.query.AppendLine("    FechaConfirmacion,");
            js.use.query.AppendLine("    Autenticada,");
            js.use.query.AppendLine("    UsuarioValidacion,");
            js.use.query.AppendLine("    FechaVerificacion,");
            js.use.query.AppendLine("    UsuarioVerificacion,");
            js.use.query.AppendLine("    EsDevolucion");
            js.use.query.AppendLine(")");
            js.use.query.AppendLine("    select");
            js.use.query.AppendLine("        getdate(),");
            js.use.query.AppendLine("        (select _name from @customer_data),");
            js.use.query.AppendLine("        (select _mail from @customer_data),");
            js.use.query.AppendLine("        _clicks,");
            js.use.query.AppendLine("        0,");
            js.use.query.AppendLine("        _name,");
            js.use.query.AppendLine("        _total,");
            js.use.query.AppendLine("        3,");
            js.use.query.AppendLine("        (select _address from @customer_data),");
            js.use.query.AppendLine("        '  NIT:   DIR: ',");
            js.use.query.AppendLine("        1,");
            js.use.query.AppendLine("        (select _phone from @customer_data),");
            js.use.query.AppendLine("        (select _number from @card_data),");
            js.use.query.AppendLine("        (select _expiry from @card_data),");
            js.use.query.AppendLine("        (select _name from @card_data),");
            js.use.query.AppendLine("        1,");
            js.use.query.AppendLine("        _total * (select ct.PorcentajeComision from CuotaTarjeta as ct where ct.CodigoRedCrediticia = 2 and ct.Cuotas = 1),");
            js.use.query.AppendLine("        0,");
            //js.use.query.AppendLine("        -- (select _code from @card_data),");
            js.use.query.AppendLine("        1,");
            js.use.query.AppendLine("        (select _type from @card_data),");
            js.use.query.AppendLine("        0,");
            js.use.query.AppendLine("        1,");
            js.use.query.AppendLine("        1,");
            js.use.query.AppendLine("        @InvoiceCode,");
            js.use.query.AppendLine("        _total,");
            js.use.query.AppendLine("        (select _code from @customer_data),");
            js.use.query.AppendLine("        1,");
            js.use.query.AppendLine("        _total,");
            js.use.query.AppendLine("        getdate(),");
            js.use.query.AppendLine("        0,");
            js.use.query.AppendLine("        0,");
            js.use.query.AppendLine("        getdate(),");
            js.use.query.AppendLine("        0,");
            js.use.query.AppendLine("        0");
            js.use.query.AppendLine("    from");
            js.use.query.AppendLine("        @advert_data");
            js.use.query.AppendLine("insert into DetalleVentaServicio");
            js.use.query.AppendLine("    select");
            js.use.query.AppendLine("        scope_identity(),");
            js.use.query.AppendLine("        5,");
            js.use.query.AppendLine("        null");

            //Insert collection
            js.use.query.AppendLine("insert into Cobro(");
            js.use.query.AppendLine("    CodigoCobro,");
            js.use.query.AppendLine("    Fecha,");
            js.use.query.AppendLine("    UsuarioCobro,");
            js.use.query.AppendLine("    NumeroTarjeta,");
            js.use.query.AppendLine("    VencimientoTarjeta,");
            js.use.query.AppendLine("    NombreTarjeta,");
            js.use.query.AppendLine("    Cuotas,");
            js.use.query.AppendLine("    MontoCuota,");
            js.use.query.AppendLine("    CodigoRedCrediticia,");
            js.use.query.AppendLine("    CodigoTipoDeTarjeta,");
            js.use.query.AppendLine("    UsuarioTarjeta,");
            js.use.query.AppendLine("    FechaConfirmacion,");
            js.use.query.AppendLine("    CodigoFactura,");
            js.use.query.AppendLine("    CodigoFormaDePago,");
            js.use.query.AppendLine("    Estado,");
            js.use.query.AppendLine("    CodigoUsuarioCreacion");
            js.use.query.AppendLine(")values(");
            js.use.query.AppendLine("    (select max(CodigoCobro) + 1 from Cobro),");
            js.use.query.AppendLine("    getdate(),");
            js.use.query.AppendLine("    0,");
            js.use.query.AppendLine("    (select _number from @card_data),");
            js.use.query.AppendLine("    (select _expiry from @card_data),");
            js.use.query.AppendLine("    (select _name from @card_data),");
            js.use.query.AppendLine("    1,");
            js.use.query.AppendLine("    (select sum(MontoCuota) from Venta where CodigoFactura = @InvoiceCode),");
            js.use.query.AppendLine("    1,");
            js.use.query.AppendLine("    (select _type from @card_data),");
            js.use.query.AppendLine("    0,");
            js.use.query.AppendLine("    getdate(),");
            js.use.query.AppendLine("    @InvoiceCode,");
            js.use.query.AppendLine("    3,");
            js.use.query.AppendLine("    1,");
            js.use.query.AppendLine("    0");
            js.use.query.AppendLine(")");

            js.use.query.AppendLine(isSimulation ? "rollback" : "commit");
            js.use.query.AppendLine("    select convert(varchar(20), @InvoiceCode) as InvoiceCode");
            js.use.query.AppendLine("end try begin catch");
            js.use.query.AppendLine("    rollback");
            js.use.query.AppendLine("    select 'Error:' + ERROR_MESSAGE()");
            js.use.query.AppendLine("end try");

            return js.execute.getValue(ref js.use.query);
        }


        private string getNumerotarjeta(string tarjeta) {
            byte[] IV = ASCIIEncoding.ASCII.GetBytes("8 @GD b!");
            byte[] EncryptionKey = Convert.FromBase64String("MTIzNDU2Nzg5MDEyMzQ1000000000000");
            byte[] buffer = Convert.FromBase64String(tarjeta);
            TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();
            tdes.Key = EncryptionKey;
            tdes.IV = IV;
            return Encoding.UTF8.GetString(tdes.CreateDecryptor().TransformFinalBlock(buffer, 0, buffer.Length));
        }

        private string sentMailCardExpire(
            string CorreoCliente,
            string tarjeta,
            ref Vars js
            ) {
            string ImagenEncabezado;
            string ImagenBarra;
            StringBuilder msg = new StringBuilder();
            string Ntarjeta;
            CorreoCliente = isSimulation ? "tec.desarrollo15.gtd@gmail.com" : CorreoCliente;

            try {
                Ntarjeta = getNumerotarjeta(tarjeta);

                js.use.Clear();
                js.use.query.AppendLine("select Valor from Parametro where CodigoParametro = 4");

                string DireccionCorreoOrigen = js.execute.getValue(ref js.use.query);
                ImagenEncabezado = "https://guatemaladigital.com:3001/images/poster_correo.jpg";
                ImagenBarra = "https://guatemaladigital.com:3001/images/CorreoEncabezado.jpg";

                msg.AppendLine("<table style='width:900px; border-width:1px;  border-style:solid'>");
                msg.AppendLine("    <tr>");
                msg.AppendLine("        <td align='center'>");
                msg.AppendLine("            <img src='" + ImagenEncabezado + "' alt='cargando imagen' width='900px'>");
                msg.AppendLine("            <img src='" + ImagenBarra + "' alt='cargando imagen' width='900px'>");
                msg.AppendLine("            <p align='center' colspan='2'>");
                msg.AppendLine("                <b>");
                msg.AppendLine("                    <h2 align='Center'> Cobro Mensual de Anuncios</h2>");
                msg.AppendLine("                </b>");
                msg.AppendLine("                <br/>");
                msg.AppendLine("                <p align='justify' colspan='2'>Estimado Sr./ Sra.</p>");
                msg.AppendLine("                <ul align='left' colspan='2'>");
                msg.AppendLine("                    <p align='justify' colspan='2'>");
                msg.AppendLine("                        Lamentamos informarle que no se realizo el cobro mensual de Anuncios, esto debido a que su tarjeta se encuentra vencida.");
                msg.AppendLine("                        <br/>");
                msg.AppendLine("                    </p>");
                msg.AppendLine("                    <li>");
                msg.AppendLine("                        <b> Tarjeta registrada: </b>");
                msg.AppendLine("                        xxxx xxxx xxxx " + Ntarjeta.Substring(Ntarjeta.Length - 4));
                msg.AppendLine("                    </li>");
                msg.AppendLine("                </ul>");
                msg.AppendLine("                <p align='justify' colspan='2'>");
                msg.AppendLine("                    Por favor, haga clic en el siguiente enlace y asocie un nueva tarjeta para el pago de mensual de Anuncios");
                msg.AppendLine("                    <a href='https://guatemaladigital.com:3001/Ingreso.aspx'>aquí</a>");
                msg.AppendLine("                    <br/>");
                msg.AppendLine("                </p>");
                msg.AppendLine("                <img src='" + ImagenBarra + "' alt='cargando imagen' width='900px'>");
                msg.AppendLine("            </p>");
                msg.AppendLine("        </td>");
                msg.AppendLine("    </tr>");
                msg.AppendLine("</table>");

                var correo = new System.Net.Mail.MailMessage();
                correo.IsBodyHtml = true;
                correo.From = new System.Net.Mail.MailAddress(DireccionCorreoOrigen, "GuatemalaDigital.com");
                correo.To.Add(CorreoCliente);
                correo.Subject = "Cobro Mensual de Anuncios";
                correo.Body = msg.ToString();
                correo.IsBodyHtml = true;
                correo.Priority = System.Net.Mail.MailPriority.Normal;
                
                string CredencialUsername = "AKIAZJ6ZLR4S2KSM7HXS";
                string CredencialPassword = "BNkV7CfzOIopvOacNiCc0l09xicWhDGBYfaRh5ozapti";

                var smtp = new System.Net.Mail.SmtpClient();
                smtp.Host = "email-smtp.us-east-1.amazonaws.com";
                smtp.Port = 587;
                smtp.EnableSsl = true;
                smtp.Credentials = new System.Net.NetworkCredential(CredencialUsername, CredencialPassword);
                smtp.Send(correo);

                return "Succesfully";
            } catch (Exception e) {
                return "Error: No se pudo enviar correo por tarjeta vencida. Correo cliente: " + CorreoCliente + " " + e.Message;
            }
        }
        private void Guardar_Datos_Archivo_Texto(string Cadena) {
            string path = "C:/inetpub/wwwroot/Sistema/CARGA/CobrosAnuncios" + DateTime.Now.Day.ToString() + DateTime.Now.Month.ToString() + ".txt";

            StreamWriter logs = !File.Exists(path) ? File.CreateText(path) : File.AppendText(path);
            logs.WriteLine(DateTime.Now.ToString());
            logs.WriteLine(Cadena);
            logs.WriteLine(" ");
        }
    }
}
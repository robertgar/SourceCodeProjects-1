using System.Text;
using System.Net.Mail;
using System.Net;
using System.Data;

namespace connection{
    public class Data {
        public Boolean isSimulation;
        public StringBuilder origin = new StringBuilder();
        public StringBuilder destination = new StringBuilder();
        public StringBuilder attachments = new StringBuilder();
        public StringBuilder subject = new StringBuilder();
        public StringBuilder body = new StringBuilder();
        public StringBuilder cc = new StringBuilder();
        public void clearAll(){
            origin.Clear();
            destination.Clear();
            attachments.Clear();
            subject.Clear();
            body.Clear();
            cc.Clear();
        }
    }
    public class SendEmails {
        public Data data = new Data();
        private commonn.UseCommon? use;
        private Execute? execute;
        public void setSimulation(ref commonn.UseCommon use, ref Execute execute) {
            this.use = use;
            this.execute = execute;
        }

        public Boolean sendEmail(ref String Error) {
            try {
                trySendEmail();
                return true;
            } catch (Exception e) {
                Error = e.ToString();
                return false;
            }
        }
        
        public void trySendEmail() {
            if (data.origin.ToString().Trim().Equals("")) {
                Console.WriteLine("Cannot send mail because Origin is empty.");
                return;
            }
            if (data.destination.ToString().Trim().Equals("")) {
                Console.WriteLine("Cannot send mail because Destination is empty");
                return;
            }
            if (data.subject.ToString().Trim().Equals("")) {
                Console.WriteLine("Cannot send mail because Title is empty.");
                return;
            }
            if (data.body.ToString().Trim().Equals("")) {
                Console.WriteLine("Cannot send mail because Body is empty.");
                return;
            }

            data.body.Replace(data.body.ToString(), connection.Properties.Resources.TemplatePrincipalMail.ToString().Replace("@body", data.body.ToString()));

            using (MailMessage mail = new MailMessage(new MailAddress(data.origin.ToString(), "GuatemalaDigital.com"), new MailAddress(data.destination.ToString()))) {
                foreach (String cc in data.cc.ToString().Split(",")) {
                    if (cc.Trim().Equals("")) { continue; }
                    mail.CC.Add(cc);
                }
                foreach (String file in data.attachments.ToString().Split(",")) {
                    if (file.Trim().Equals("")) { continue; }
                    mail.Attachments.Add(new Attachment(file));
                }

                mail.Subject = data.subject.ToString();
                mail.Body = data.body.ToString();
                mail.IsBodyHtml = true;
                mail.Priority = MailPriority.Normal;
                using (SmtpClient sender = new SmtpClient()) {
                    sender.UseDefaultCredentials = false;
                    sender.Host = "email-smtp.us-east-1.amazonaws.com";
                    sender.Port = 587;
                    sender.EnableSsl = true;
                    sender.Credentials = new NetworkCredential("AKIAZJ6ZLR4SSIO3C2X6", "BCuKYZn5qkcXIFT5YNSxgVu0uQejzt+w4q8+F6hUIROW");
                    sender.Send(mail);
                }
            }
        }

        public void sendMailTracking(ref String Tracking, ref StringBuilder query, ref Execute execute) {
            query.Clear();
            query.AppendLine(" select distinct");
            query.AppendLine("    v.CodigoVenta as SalesCode,");
            query.AppendLine("    convert(varchar, v.FechaConfirmacion, 103) as Date,");
            query.AppendLine("    v.NombreProducto as Product,");
            query.AppendLine("    v.NombreCliente as Applicant,");
            query.AppendLine("    v.CorreoCliente as Mail,");
            query.AppendLine("    v.Telefonos,");
            query.AppendLine("    concat(v.DireccionDeEntrega, iif(m.Nombre is null, '', ' Municipio:'), m.Nombre, iif(d.Nombre is null, '', ' Departamento:'), d.Nombre) as Address,");
            query.AppendLine("    fp.Nombre as PaymentMethod,");
            query.AppendLine("    v.Monto as Price,");
            query.AppendLine("    iif(");
            query.AppendLine("        v.CodigoFormaDePago = 3,");
            query.AppendLine("        (isnull(v.Monto, 0) + isnull(f.Envio, 0) + isnull(t.TotalServicios, 0))*(select PorcentajeComision from CuotaTarjeta where CodigoRedCrediticia = 2 and Cuotas = isnull(v.Cuotas, 1))/ isnull(v.Cuotas, 1),");
            query.AppendLine("        isnull(v.Monto, 0) + isnull(f.Envio, 0) + isnull(t.TotalServicios, 0)");
            query.AppendLine("    ) as AmountPayable,");
            query.AppendLine("    (select Valor from Parametro where CodigoParametro = 8) as Days,");
            query.AppendLine("    trim(");
            query.AppendLine("        iif(");
            query.AppendLine("            charindex('NIT:', v.Factura) > 0,");
            query.AppendLine("            substring(v.Factura, 0, charindex('NIT:', v.Factura)),");
            query.AppendLine("            iif(");
            query.AppendLine("                charindex('DIR:', v.Factura) >0,");
            query.AppendLine("                substring(v.Factura, 0, charindex('DIR:', v.Factura)),");
            query.AppendLine("                v.Factura");
            query.AppendLine("            )");
            query.AppendLine("        )");
            query.AppendLine("    ) as Name,");
            query.AppendLine("    trim(");
            query.AppendLine("        iif(");
            query.AppendLine("            charindex('NIT:', v.Factura) = 0,");
            query.AppendLine("            '',");
            query.AppendLine("            iif(");
            query.AppendLine("                charindex('DIR:', v.Factura) > 0,");
            query.AppendLine("                substring(v.Factura, charindex('NIT:', v.Factura) + 4, charindex('DIR:', v.Factura) - charindex('NIT:', v.Factura) - 4),");
            query.AppendLine("                substring(v.Factura, charindex('NIT:', v.Factura) + 4, len(v.Factura) - charindex('NIT:', v.Factura))");
            query.AppendLine("            )");
            query.AppendLine("        )");
            query.AppendLine("    ) as Nit,");
            query.AppendLine("    trim(");
            query.AppendLine("        iif(");
            query.AppendLine("            charindex('DIR:', v.Factura) > 0,");
            query.AppendLine("            substring(v.Factura, charindex('DIR:', v.Factura) + 4, len(v.Factura) - charindex('DIR:', v.Factura) - 4),");
            query.AppendLine("            ''");
            query.AppendLine("        )");
            query.AppendLine("    ) as InvoiceAddress,");
            query.AppendLine("    pr.Foto as UrlImage,");
            query.AppendLine("    replace(ee.Nombre, 'Bodega', 'en bodega') as DeliveryStatus,");
            query.AppendLine("    concat('https://guatemaladigital.com/Producto/', pr.CodigoProducto) as UrlProduct");
            query.AppendLine(" from");
            query.AppendLine("    VentaPedido as vp");
            query.AppendLine("    inner join Pedido  as pe on pe.CodigoPedido = vp.CodigoPedido");
            query.AppendLine("    inner join Venta as v");
            query.AppendLine("        on v.CodigoVenta = vp.CodigoVenta");
            query.AppendLine("        and v.CodigoEstadoEntrega = 2");
            query.AppendLine("        and v.CodigoEstadoDeVenta = 1");
            query.AppendLine("    inner join Paquete as pa");
            query.AppendLine("        on pa.CodigoPaquete = pe.CodigoPaquete");
            query.AppendLine("        and pa.GuiaAerea is not null");
            query.AppendLine("    inner join FormaDePago as fp on fp.CodigoFormaDePago = v.CodigoFormaDePago");
            query.AppendLine("    inner join Factura as f on f.CodigoFactura = v.CodigoFactura");
            query.AppendLine("    inner join Producto as pr on pr.CodigoProducto = v.CodigoProducto");
            query.AppendLine("    inner join EstadoEntrega as ee on ee.CodigoEstadoEntrega = v.CodigoEstadoEntrega");
            query.AppendLine("    left join Departamento as d on d.CodigoDepartamento = v.CodigoDepartamento");
            query.AppendLine("    left join Municipio as m on m.CodigoMunicipio = v.CodigoMunicipio");
            query.AppendLine("    left join (");
            query.AppendLine("        select");
            query.AppendLine("            dvs.CodigoVenta,");
            query.AppendLine("            sum(dvs.Monto) as TotalServicios");
            query.AppendLine("        from");
            query.AppendLine("            DetalleVentaServicio as dvs");
            query.AppendLine("        where");
            query.AppendLine("            dvs.CodigoServicio in (2, 3, 4)");
            query.AppendLine("        group by");
            query.AppendLine("            dvs.CodigoVenta");
            query.AppendLine("    ) as t on t.CodigoVenta = v.CodigoVenta");
            query.AppendLine(" where");
            query.Append("    pa.CodigoDeRastreo like '%").Append(Tracking).Append("%'");

            DataTable table = new DataTable();
            execute.fillTable(ref query, ref table);

            if (table.Rows.Count == 0) { return; }
            DataTable tablita = new DataTable();
            StringBuilder ContentStatus = new StringBuilder();
            String Status = "";

            foreach (DataRow row in table.Rows) {
                query.Clear();
                query.Append(" declare @Tracking varchar(30) = '%").Append(Tracking).AppendLine("%'");
                query.AppendLine(" select");
                query.AppendLine("    substring(sp.Value, 1, 1) as Id,");
                query.AppendLine("    substring(sp.Value, 3, len(sp.Value)) as Date,");
                query.AppendLine("    ee.Etapa,");
                query.AppendLine("    replace(ee.Nombre, 'Bodega', 'En bodega') as Nombre");
                query.AppendLine(" from");
                query.AppendLine("    string_split(");
                query.AppendLine("        (");
                query.AppendLine("            select top 1");
                query.AppendLine("                concat(';1-', convert(varchar, pe.FechaDeOrden, 103)) +");
                query.AppendLine("                concat(';2-', convert(varchar, pe.FechaDeEnvio, 103)) +");
                query.AppendLine("                concat(';3-', convert(varchar, pa.Fecha, 103)) +");
                query.AppendLine("                concat(';4-', convert(varchar, iif(pa.FechaRecibida is null, v.FechaConfirmacion, pa.FechaRecibida), 103)) +");
                query.AppendLine("                concat(';5-', convert(varchar, f.Fecha, 103)) +");
                query.AppendLine("                concat(';6-', convert(varchar, v.FechaEntrega, 103)) as Dates");
                query.AppendLine("            from");
                query.AppendLine("                VentaPedido as vp");
                query.AppendLine("                inner join Pedido as pe on pe.CodigoPedido = vp.CodigoPedido");
                query.AppendLine("                inner join Paquete as pa on pa.CodigoPaquete = pe.CodigoPaquete");
                query.AppendLine("                inner join Venta as v on v.CodigoVenta = vp.CodigoVenta");
                query.AppendLine("                inner join Factura as f on f.CodigoFactura = v.CodigoFactura");
                query.AppendLine("            where");
                query.AppendLine("                pa.CodigoDeRastreo like @Tracking");
                query.AppendLine("                and v.CodigoEstadoEntrega = 2");
                query.AppendLine("                and v.CodigoEstadoDeVenta =  1");
                query.AppendLine("                and pa.GuiaAerea is not null");
                query.AppendLine("        ), ';'");
                query.AppendLine("    ) as sp");
                query.AppendLine("    left join EstadoEntrega as ee on ee.CodigoEstadoEntrega = iif(charindex('-', sp.Value) > 0, substring(sp.Value, 1, 1), '')");
                query.AppendLine(" where");
                query.AppendLine("    len(sp.Value) > 2");

                execute.fillTable(ref query, ref tablita);

                if (tablita.Rows.Count == 0) { continue; }
                ContentStatus.Clear();

                foreach (DataRow fila in tablita.Rows) {
                    ContentStatus.AppendLine("<td>");
                    ContentStatus.AppendLine(fila["Nombre"].ToString());
                    ContentStatus.AppendLine("<br/>");
                    ContentStatus.AppendLine(fila["Date"].ToString());
                    ContentStatus.AppendLine("</td>");
                    Status = fila["Etapa"].ToString();
                }

                if (ContentStatus.ToString().Trim().Equals("")) { continue; }

                data.clearAll();
                data.body.Append(Properties.Resources.TemplateEmailTracking.ToString());
                data.body.Replace("@V1", tablita.Rows.Count.ToString());
                data.body.Replace("@ContentStatus", ContentStatus.ToString());
                data.body.Replace("@SalesCode", row["SalesCode"].ToString());
                data.body.Replace("@Status", Status);
                data.body.Replace("@Product", row["Product"].ToString());
                data.body.Replace("@Applicant", row["Applicant"].ToString());
                data.body.Replace("@Mail", row["Mail"].ToString());
                data.body.Replace("@Date", row["Date"].ToString());
                Status = row["Telefonos"].ToString();
                if (Status.Contains("-")) {
                    data.body.Replace("@Phone", Status.Substring(0, Status.IndexOf("-")));
                    data.body.Replace("@AlternatePhone", Status.Substring(Status.IndexOf("-")));
                }
                if (Status.Contains("||")) {
                    data.body.Replace("@Phone", Status.Substring(0, Status.IndexOf("||")));
                    data.body.Replace("@AlternatePhone", Status.Substring(Status.IndexOf("||")));
                }
                data.body.Replace("@Phone", Status);
                data.body.Replace("@AlternatePhone", "");
                data.body.Replace("@Address", row["Address"].ToString());
                data.body.Replace("@PaymentMethod", row["PaymentMethod"].ToString());
                data.body.Replace("@Price", "Q " + row["Price"].ToString());
                data.body.Replace("@AmountPayable", "Q " + row["AmountPayable"].ToString());
                data.body.Replace("@Days", row["Days"].ToString());
                data.body.Replace("@Name", row["Name"].ToString());
                data.body.Replace("@Nit", row["Nit"].ToString());
                data.body.Replace("@InvoiceAddress", row["InvoiceAddress"].ToString());
                data.body.Replace("@UrlImage", row["UrlImage"].ToString());
                data.body.Replace("@UrlProduct", row["UrlProduct"].ToString());

                data.subject.Append("Your GuatemalaDigital ").Append(row["SalesCode"]);
                data.subject.Append(" order is ").Append(row["DeliveryStatus"]);
                data.origin.Append(execute.getParameter(123));
                
                if (data.isSimulation) {
                    data.destination.Append("tec.desarrollo15.gtd@gmail.com");
                }else{
                    data.destination.Append(row["Mail"]);
                }

                sendEmail(ref Status);
            }
        }
    }
}

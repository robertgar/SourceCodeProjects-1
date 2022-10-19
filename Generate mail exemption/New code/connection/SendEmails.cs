using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using connection;
using System.Data;
using common;

namespace connection{
    public class Data {
        public Boolean isSimulation;
        public String? origin;
        public String? destination;
        public StringBuilder title = new StringBuilder();
        public StringBuilder body = new StringBuilder();
        public StringBuilder buffer = new StringBuilder();
    }
    public class SendEmails {
        private Data? data;
        private UseCommon? use;
        private Execute? execute;
        public void setSimulation(ref Data data, ref UseCommon use, ref Execute execute) {
            this.data = data;
            this.use = use;
            this.execute = execute;
        }

        private StringBuilder structure = new StringBuilder();
        public void sendEmail() {
            if (data.isSimulation) {
                Console.WriteLine("Email has been sent successfully!");
                return;
            }

            if (String.IsNullOrEmpty(data.origin)) {
                Console.WriteLine("Cannon send mail because Origin is empty.");
                return;
            }
            if (String.IsNullOrEmpty(data.destination)) {
                Console.WriteLine("Cannot send mail because Destination is empty");
                return;
            }
            if (data.title.Length == 0) {
                Console.WriteLine("Cannot send mail because Title is empty.");
                return;
            }
            if (data.body.Length == 0) {
                Console.WriteLine("Cannot send mail because Body is empty.");
                return;
            }

            data.buffer.Clear();
            data.buffer.AppendLine("<table style = ccwidth:800px; border-width:1px;  border-style:solidcc>");
            data.buffer.AppendLine("    <tr>");
            data.buffer.AppendLine("        <td>");
            data.buffer.AppendLine("            <img src=cchttps://guatemaladigital.com:3001/images/poster_correo.jpgcc alt=cccargando imagen...cc width=cc800pxcc>");
            data.buffer.AppendLine("            <img src=cchttps://guatemaladigital.com:3001/images/CorreoEncabezado.jpgcc alt=cccargando imagen...cc width=cc800pxcc>");
            data.buffer.AppendLine(data.body.ToString());
            data.buffer.AppendLine("            <img src=cchttps://guatemaladigital.com:3001/images/CorreoEncabezado.jpgcc alt=cccargando imagen...cc width=cc800pxcc>");
            data.buffer.AppendLine("        </td>");
            data.buffer.AppendLine("    </tr>");
            data.buffer.AppendLine("</table>");

            MailMessage Email = new MailMessage();
            Email.From = new MailAddress(data.origin, "GuatemalaDigital.com");
            Email.To.Add(data.destination);
            Email.Subject = data.title.ToString();
            Email.Body = data.buffer.ToString();
            Email.IsBodyHtml = true;
            Email.Priority = MailPriority.Normal;
            SmtpClient sender = new SmtpClient();
            sender.EnableSsl = true;
            sender.UseDefaultCredentials = false;
            sender.Host = "email-smtp.us-east-1.amazonaws.com";
            sender.Port = 587;
            sender.Credentials = new System.Net.NetworkCredential("AKIAZJ6ZLR4S6FRFMBM6", "BP42Ou4Xqtc758pmTKyhFBuRnTZxxx4pOPWO8frYpD8F");

            try {
                sender.Send(Email);
            } catch (Exception e) {
                Console.WriteLine(e);
            }
        }

        public void sendTrackingEmail(ref String Tracking, ref Data dat) {
            use.query.AppendLine(" Declare @Sale int = (");
            use.query.AppendLine("    select");
            use.query.AppendLine("        max(pa.CodigoPaquete)");
            use.query.AppendLine("    from");
            use.query.AppendLine("        Paquete as pa");
            use.query.AppendLine("    where");
            use.query.Append("        pa.CodigoDeRastreo like '%").Append(Tracking).Append("%'");
            use.query.AppendLine(" )");
            use.query.AppendLine(" select");
            use.query.AppendLine("    v.CodigoVenta,");
            use.query.AppendLine("    v.CodigoProducto,");
            use.query.AppendLine("    v.FechaConfirmacion,");
            use.query.AppendLine("    convert(varchar, v.Cantidad) + ' ' + v.NombreProducto as NombreProducto,");
            use.query.AppendLine("    v.NombreCliente,");
            use.query.AppendLine("    v.CorreoCliente,");
            use.query.AppendLine("    v.Telefonos,");
            use.query.AppendLine("    isnull(convert(int, v.EsPedido), 0) as EsPedido,");
            use.query.AppendLine("    v.DireccionDeEntrega + ' Departamento: ' + d.Nombre + ' Municipio: ' + m.Nombre as DireccionDeEntrega,");
            use.query.AppendLine("    v.CodigoFormaDePago,");
            use.query.AppendLine("    fp.Nombre as FormaDePago,");
            use.query.AppendLine("    v.Monto,");
            use.query.AppendLine("    v.Envio,");
            use.query.AppendLine("    isnull(v.Cuotas, 1) as Cuotas,");
            use.query.AppendLine("    (");
            use.query.AppendLine("        case v.CodigoFormaDePago");
            use.query.AppendLine("            when 1 then v.ServicioPagoEfectivo");
            use.query.AppendLine("            when 3 then v.MontoCuota - v.Monto");
            use.query.AppendLine("        else");
            use.query.AppendLine("            0.00");
            use.query.AppendLine("        end");
            use.query.AppendLine("    ) as ServicioPago,");
            use.query.AppendLine("    (");
            use.query.AppendLine("        v.Monto + v.Envio + (");
            use.query.AppendLine("            case v.CodigoFormaDePago");
            use.query.AppendLine("                when 1 then v.ServicioPagoEfectivo");
            use.query.AppendLine("                when 3 then v.MontoCuota - v.Monto");
            use.query.AppendLine("            else");
            use.query.AppendLine("                0.00");
            use.query.AppendLine("            end");
            use.query.AppendLine("        )");
            use.query.AppendLine("    ) as Total,");
            use.query.AppendLine("    v.Factura,");
            use.query.AppendLine("    p.Foto,");
            use.query.AppendLine("    v.NombreRecibido,");
            use.query.AppendLine("    v.FechaEntrega,");
            use.query.AppendLine("    ee.Nombre as EmpresaDeEntrega,");
            use.query.AppendLine("    v.CodigoEstadoEntrega");
            use.query.AppendLine(" from");
            use.query.AppendLine("    Venta as v");
            use.query.AppendLine("    inner join FormaDePago as fp on fp.CodigoFormaDePago = v.CodigoFormaDePago");
            use.query.AppendLine("    inner join Producto as p on p.CodigoProducto = v.CodigoProducto");
            use.query.AppendLine("    left join EmpresaDeEntrega as ee on ee.CodigoEmpresaDeEntrega = v.CodigoEmpresaDeEntrega");
            use.query.AppendLine("    left join Departamento as d on d.CodigoDepartamento = v.CodigoDepartamento");
            use.query.AppendLine("    left join Municipio as m on m.CodigoMunicipio = v.CodigoMunicipio");
            use.query.AppendLine(" where");
            use.query.AppendLine("    v.CodigoVenta = @Sale");
            use.query.AppendLine("    and v.CodigoEstadoDeVenta = 1");

            execute.fillTable(ref use.query, ref use.tabBuffer);

            if (use.tabBuffer.Rows.Count == 0) { return; }

            dataEmail data = new dataEmail();
            data.origin = execute.getParameter(123);

            foreach (DataRow row in use.tabBuffer.Rows) {
                data.destination = row["CorreoCliente"].ToString();
                data.SaleCode = row["CodigoVenta"].ToString();

                if (row["CodigoEstadoEntrega"].ToString().Equals("6")){
                    data.BackColor = "#218748";
                    data.CssTrackinSpan = "delivery-status-text status-green";
                }else{
                    data.BackColor = "#f98c31";
                    data.CssTrackinSpan = "delivery-status-text status-orange";
                }

                data.title.Clear();
                data.title.Append("Tu orden: ").Append(data.SaleCode).Append(", de GuatemalaDigital.com, está en estado: ").Append(row["CodigoEstadoEntrega"].ToString());
            }

            use.query.Clear();
            use.query.AppendLine(" select");
            use.query.AppendLine("    case");
            use.query.AppendLine("        when v.EsPedido IS Not null then 6");
            use.query.AppendLine("        when v.EsPedido IS NULL and f.CodigoDeRastreo IS not null then 3");
            use.query.AppendLine("        when v.EsPedido IS NULL and f.CodigoDeRastreo IS null then 2");
            use.query.AppendLine("    end as Pasos");
            use.query.AppendLine(" from");
            use.query.AppendLine("    Venta v");
            use.query.AppendLine("    Left Join Factura f on v.CodigoFactura = f.CodigoFactura");
            use.query.AppendLine(" where");
            use.query.AppendLine("    CodigoVenta = ").Append(data.SaleCode);
            data.Steps = execute.getNat(ref use.query);

            use.query.Clear();
            use.query.Append(" declare @SalesCode int = ").Append(data.SaleCode);
            use.query.AppendLine(" select");
            use.query.AppendLine("    ee.CodigoEstadoEntrega,");
            use.query.AppendLine("    ee.Nombre,");
            use.query.AppendLine("    ee.Etapa,");
            use.query.AppendLine("    dt.Datita");
            use.query.AppendLine(" from");
            use.query.AppendLine("    EstadoEntrega as ee");
            use.query.AppendLine("    inner join (");
            use.query.AppendLine("        select");
            use.query.AppendLine("            substring(sp.Value, 1, charindex(':', sp.Value) - 1) as DeliveryStatusCode,");
            use.query.AppendLine("            substring(sp.Value, charindex(':', sp.Value) + 1, len(sp.Value)) as Datita");
            use.query.AppendLine("        from");
            use.query.AppendLine("            string_split(");
            use.query.AppendLine("                (");
            use.query.AppendLine("                    select");
            use.query.AppendLine("                        iif(pe.FechaDeOrden is not null, ',1:' + convert(varchar, pe.FechaDeOrden, 103), '') +");
            use.query.AppendLine("                        iif(pe.FechaDeEnvio is not null, ',2:' + convert(varchar, pe.FechaDeEnvio, 103), '') +");
            use.query.AppendLine("                        iif(pa.Fecha is not null, ',3:' + convert(varchar, pa.Fecha, 103), '') +");
            use.query.AppendLine("                        iif(");
            use.query.AppendLine("                            iif(pa.CodigoPaquete = pe.CodigoPaquete and pa.FechaRecibida is not null, pa.FechaRecibida, v.FechaConfirmacion) is not null, ");
            use.query.AppendLine("                                ',4:' + convert(varchar, iif(pa.CodigoPaquete = pe.CodigoPaquete and pa.FechaRecibida is not null, pa.FechaRecibida, v.FechaConfirmacion), 103)");
            use.query.AppendLine("                                , ''");
            use.query.AppendLine("                        ) +");
            use.query.AppendLine("                        iif(f.Fecha is not null, ',5:'+convert(varchar, f.Fecha, 103), '') +");
            use.query.AppendLine("                        iif(v.FechaEntrega is not null, ',6:' + convert(varchar, v.FechaEntrega, 103), '') as Result");
            use.query.AppendLine("                    from");
            use.query.AppendLine("                        Venta as v");
            use.query.AppendLine("                        inner join VentaPedido as vp on vp.CodigoVenta = v.CodigoVenta");
            use.query.AppendLine("                        inner join Pedido as pe on pe.CodigoPedido = vp.CodigoPedido");
            use.query.AppendLine("                        left join Paquete as pa on pa.CodigoPaquete = pe.CodigoPaquete");
            use.query.AppendLine("                        inner join Factura as f on f.CodigoFactura = v.CodigoFactura");
            use.query.AppendLine("                    where");
            use.query.AppendLine("                        v.CodigoVenta = @SalesCode");
            use.query.AppendLine("                ), ','");
            use.query.AppendLine("            ) as sp");
            use.query.AppendLine("        where");
            use.query.AppendLine("            trim(sp.Value) != ''");
            use.query.AppendLine("    ) as dt on dt.DeliveryStatusCode = ee.CodigoEstadoEntrega");

            execute.fillTable(ref use.query, ref use.tabBuffer);

            if (use.tabBuffer.Rows.Count == 0) { return; }

            foreach (DataRow row in use.tabBuffer.Rows) {
                data.status = row["Etapa"].ToString();
            }

            data.body.Clear();
            data.body.AppendLine("<table>");
            data.body.AppendLine("    <tr>");
            data.body.AppendLine("        <td align=cccentercc>");
            data.body.AppendLine("            <p align=cccentercc colspan=cc2cc>");
            data.body.Append("                <b> Número de orden:").Append(data.SaleCode).Append("</b>");
            data.body.AppendLine("            </p>");
            data.body.AppendLine("            <table>");
            data.body.AppendLine("                <tr>");
            data.body.AppendLine("                    <td>");
            data.body.AppendLine("                        <p>");
            data.body.AppendLine("                            <b>");
            data.body.Append("                                <font color=cc").Append(data.BackColor).Append("cc>").Append(data.title.ToString()).Append("</font>");
            data.body.AppendLine("                            </b>");
            data.body.AppendLine("                        </p>");
            data.body.AppendLine("                    </td>");
            data.body.AppendLine("                </tr>");
            data.body.AppendLine("            </table>");
            data.body.AppendLine("        </br>");
            data.body.AppendLine("        <table width=cc90%cc border = cc1cc>");
            data.body.AppendLine("            <tr>");
            data.body.AppendLine("                <td colspan=cc2cc align=cccentercc>");
            data.body.AppendLine("                    <div id=ccDivTrackingRastreocc class=cctracking-graphic-columncc style=ccmargin-bottom: 0px; border: 3px solid grey;cc>");
            data.body.AppendLine("                        <div style=ccmargin-top: 3%; margin-bottom: 1%cc>");
            data.body.Append("                            <span id=ccSpanStatusRastreocc class=cc").Append(data.CssTrackinSpan).Append("cc style=ccwidth:40%cc>");
            data.body.Append("                                ").Append(data.status);
            data.body.AppendLine("                            </span>");
            data.body.AppendLine("                        </div>");
            data.body.AppendLine("                        <div class=cctrack-bar-holdercc>");
            data.body.AppendLine("                            <div id=ccDivTrackStatusBarcc class=cctrack-status-barcc>");

            for (int i = 1; i < data.Steps; i++) {

            }

            sendEmail();
        }
    }

    internal class dataEmail{
        public StringBuilder title = new StringBuilder();
        public String destination;
        public String origin;
        public String SaleCode;
        public String status;
        public String BackColor;
        public String CssTrackinSpan;
        public int Steps;
        public StringBuilder body = new StringBuilder();
    }
}

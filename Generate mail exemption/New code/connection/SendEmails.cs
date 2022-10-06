using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using connection;
using System.Data;

namespace connection{
    public class SendEmails {
        private Boolean isSimulation;
        private StringBuilder body = new StringBuilder();
        private common.UseCommon use = new common.UseCommon();
        private Execute execute = new Execute();

        public void setSimulation(ref Boolean isSimulation) {
            this.isSimulation = isSimulation;
        }

        private StringBuilder structure = new StringBuilder();
        public void sendEmail(ref String origin, ref String destination, ref String title, ref StringBuilder msg) {
            if (isSimulation) {
                Console.WriteLine("Email has been sent successfully!");
                return;
            }

            body.Clear();
            body.AppendLine("<table style = ccwidth:800px; border-width:1px;  border-style:solidcc>");
            body.AppendLine("    <tr>");
            body.AppendLine("        <td>");
            body.AppendLine("            <img src=cchttps://guatemaladigital.com:3001/images/poster_correo.jpgcc alt=cccargando imagen...cc width=cc800pxcc>");
            body.AppendLine("            <img src=cchttps://guatemaladigital.com:3001/images/CorreoEncabezado.jpgcc alt=cccargando imagen...cc width=cc800pxcc>");
            body.AppendLine(msg.ToString());
            body.AppendLine("            <img src=cchttps://guatemaladigital.com:3001/images/CorreoEncabezado.jpgcc alt=cccargando imagen...cc width=cc800pxcc>");
            body.AppendLine("        </td>");
            body.AppendLine("    </tr>");
            body.AppendLine("</table>");

            MailMessage Email = new MailMessage();
            Email.From = new MailAddress(origin, "GuatemalaDigital.com");
            Email.To.Add(destination);
            Email.Subject = title;
            Email.Body = body.ToString();
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

        public void sendTrackingEmail(ref String Tracking) {
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
            use.query.AppendLine("    ee.Nombre as EmpresaDeEntrega");
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

            dataEmail data = new dataEmail();
            data.origin = execute.getParameter(123);

            foreach (DataRow row in use.tabBuffer.Rows) {
                data.destination = row["CorreoCliente"].ToString();
                data.SaleCode = row["CodigoVenta"].ToString();
            }

            data.title = "Tu orden: " + data.SaleCode + ", de GuatemalaDigital.com, está en estado: " + data.status;
            sendEmail(ref data.origin, ref data.destination, ref data.destination, ref data.body);
        }
    }

    internal class dataEmail{
        public String title;
        public String destination;
        public String origin;
        public String SaleCode;
        public String status;
        public StringBuilder body = new StringBuilder();
    }
}

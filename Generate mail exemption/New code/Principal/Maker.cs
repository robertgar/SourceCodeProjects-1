using System.Text;
using System.Data;

namespace Principal {
    internal class Maker {
        private common.UseCommon use = new common.UseCommon();
        private connection.Execute execute = new connection.Execute();
        private connection.SendEmails email = new connection.SendEmails();
        private common.GuidePdf pdf = new common.GuidePdf();
        private connection.Slack slack = new connection.Slack();
        private int CounterGeneral;

        public Maker(Boolean DeAMentis) {
            email.data.isSimulation = DeAMentis;
            execute.setSimulation(ref DeAMentis);
            email.setSimulation(ref use, ref execute);
            CounterGeneral = 0;
        }
        
        public void makeAll(DateTime tiempito) {
            slack.data.Clear();
            slack.data.alertTitle.Append("Alert");
            slack.data.subject.Append("Procedure alert info: the process has begun...");
            slack.data.message.Append("The process 'generate mail exemption' has begun.");
            slack.data.message.Append("\nTime start: ").Append(tiempito);
            slack.data.warningColour = slack.data.selectColour.Green;
            slack.send();

            try {
                tryMakeAll();
            } catch (Exception e) {
                Console.WriteLine(e);
                slack.data.Clear();
                slack.data.alertTitle.Append("Error");
                slack.data.subject.Append("Procedure alert inf: Ops! Something has gone wrong. :(");
                slack.data.message.Append(e);
                slack.send();
            }

            slack.data.Clear();
            slack.data.alertTitle.Append("Alert");
            slack.data.subject.Append("Procedure alert info: the process has been completed!");
            slack.data.message.Append("The process 'generate mail exemption' has been completed.");
            slack.data.message.Append("\nTime execution total: ").Append(DateTime.Now - tiempito);
            slack.data.message.Append("\nTotal emails sent: ").Append(CounterGeneral);
            slack.data.warningColour = slack.data.selectColour.Green;
            slack.send();
        }
        private void tryMakeAll() {
            use.query.Clear();
            use.query.AppendLine(" declare @Temp table(AmazonOrder varchar(50))");
            use.query.AppendLine(" insert into @Temp");
            use.query.AppendLine("    select");
            use.query.AppendLine("        oc.OrdenDeAmazon");
            use.query.AppendLine("    from");
            use.query.AppendLine("        OrdenDeCompra as oc");
            use.query.AppendLine("        inner join Pedido as p on p.OrdenDeAmazon = oc.OrdenDeAmazon");
            use.query.AppendLine("    where");
            use.query.AppendLine("        oc.CorreoDeExencion is null");
            use.query.AppendLine("    group by");
            use.query.AppendLine("        oc.OrdenDeAmazon");
            use.query.AppendLine("    having");
            use.query.AppendLine("        isnull(");
            use.query.AppendLine("            sum(");
            use.query.AppendLine("                iif(p.Cantidad > 0 and p.CodigoEstadoPedido in (1, 2, 6), 1, 0)");
            use.query.AppendLine("            ), 0");
            use.query.AppendLine("        ) = 0");

            use.query.AppendLine(" select top 3");
            use.query.AppendLine("    pe.OrdenDeAmazon as AmazonOrder,");
            use.query.AppendLine("    isnull(");
            use.query.AppendLine("        sum(");
            use.query.AppendLine("            iif(pe.CodigoEstadoPedido = 3, 1, 0)");
            use.query.AppendLine("        ), 0");
            use.query.AppendLine("    ) as ProductsReceived");
            use.query.AppendLine(" from");
            use.query.AppendLine("    Pedido as pe");
            use.query.AppendLine("    inner join CuentaDeCompra as cc on cc.Correo = pe.Correo");
            use.query.AppendLine("    inner join @Temp as t on t.AmazonOrder = pe.OrdenDeAmazon");
            use.query.AppendLine(" group by");
            use.query.AppendLine("    pe.OrdenDeAmazon");
            use.query.AppendLine(" having");
            use.query.AppendLine("    (");
            use.query.AppendLine("        isnull(SUM(iif(pe.CodigoDeRastreo is not null, pe.Impuesto, 0)),0) > 0");
            use.query.AppendLine("        or isnull(sum(iif(cc.SiempreEnviarExencion = 1, 1, 0)), 0) > 0");
            use.query.AppendLine("    )");
            use.query.AppendLine("    and isnull(sum(iif(cc.NuncaEnviarExencion = 1, 1, 0)), 0) = 0");
            use.query.AppendLine("    and isnull(sum(iif(pe.Cantidad > 0 and pe.CodigoEstadoPedido != 4, 1, 0)), 0) > 0");

            DataTable tabMain = new DataTable();

            Console.WriteLine("Filling tab...");
            execute.fillTable(ref use.query, ref tabMain);
            Console.WriteLine("Sending emails...");

            if (tabMain.Rows.Count == 0) { return; }

            StringBuilder GuidesList = new StringBuilder();
            StringBuilder attachments = new StringBuilder();
            DataTable tablita = new DataTable();
            String Error = "";

            foreach (DataRow row in tabMain.Rows) {
                generateGuide(row, ref tablita);
                if (!verifyPackage(row["AmazonOrder"].ToString(), ref use.Text, ref GuidesList, ref tablita)) { continue; }

                use.query.Clear();
                use.query.AppendLine(" select distinct");
                use.query.AppendLine("    pa.GuiaAerea,");
                use.query.AppendLine("    pa.CodigoPaquete,");
                use.query.AppendLine("    pa.EnviadoPor,");
                use.query.AppendLine("    pa.Descripcion,");
                use.query.AppendLine("    iif(isnull(pa.Peso, 0) != 0, convert(varchar, pa.Peso) + ' punds', convert(varchar, pa.Peso)) as Peso,");
                use.query.AppendLine("    iif(isnull(pa.Peso, 0) != 0, '$' + convert(varchar, pa.Peso * 1.6), 'Null') as Charge,");
                use.query.AppendLine("    (select sum(p.Cantidad) from Pedido  as p where p.CodigoPaquete = pe.CodigoPaquete) as Parts,");
                use.query.AppendLine("    pa.CodigoDeRastreo,");
                use.query.AppendLine("    iif(ISNULL(pe.CodigoImportador,0) > 0, pe.CodigoImportador, isnull(pa.CodigoImportador,0)) CodigoImportador,");
                use.query.AppendLine("    i.Nombre,");
                use.query.AppendLine("    convert(varchar, pe.FechaDeEnvio, 103) as FechaEnvio,");
                use.query.AppendLine("    i.Linea1AWB + iif(isnull(i.Linea2AWB, '') != '', '<br />' + i.Linea2AWB, '') + iif(isnull(i.Linea3AWB, '') != '', '<br />' + i.Linea3AWB, '')as Awb,");
                use.query.AppendLine("    iif(charindex('NO-Enviar-Exencion', replace(pe.Observaciones, 'ó', 'o')) > 0, 0, 1) as EnviarExencion");
                use.query.AppendLine(" from");
                use.query.AppendLine("    Pedido pe");
                use.query.AppendLine("    inner join Paquete pa on pa.CodigoPaquete = pe.CodigoPaquete");
                use.query.AppendLine("    inner join Importador i on i.CodigoImportador = pe.CodigoImportador");
                use.query.AppendLine("    inner join (");
                use.query.AppendLine("        select");
                use.query.AppendLine("            trim(sp.Value) as Guide");
                use.query.AppendLine("        from");
                use.query.Append("            string_split('").Append(GuidesList).AppendLine("', ',') as sp");
                use.query.AppendLine("        where");
                use.query.AppendLine("            trim(sp.Value) != ''");
                use.query.AppendLine("    ) as t on t.Guide = pa.GuiaAerea");
                use.query.AppendLine(" where");
                use.query.AppendLine("    pe.Cantidad > 0");
                use.query.AppendLine("    And pe.CodigoEstadoPedido != 4");

                execute.fillTable(ref use.query, ref tablita);

                if (tablita.Rows.Count == 0) { continue; }

                use.Flag = true;
                foreach (DataRow RowGuide in tablita.Rows) {
                    if (!use.Flag) { continue; }

                    //Create PDF report
                    use.Flag = pdf.createPDFReport(RowGuide, ref attachments, ref Error);
                }
                GuidesList.Clear();

                if (!use.Flag) {
                    slack.data.Clear();
                    slack.data.alertTitle.Append("Error");
                    slack.data.subject.Append("Error procedure: create PDF report, amazon order: ").Append(row["AmazonOrder"]).Append("");
                    slack.data.message.Append(Error);
                    slack.data.warningColour = slack.data.selectColour.Red;
                    slack.send();
                    continue;
                }

                use.query.Clear();
                use.query.Append(" declare @AmazonOrder varchar(50) = '").Append(row["AmazonOrder"].ToString()).AppendLine("'");
                use.query.AppendLine(" select");
                use.query.AppendLine("    pe.Correo,");
                use.query.AppendLine("    cc.EnviarCopiaExencion,");
                use.query.AppendLine("    count(distinct pe.CodigoPedido) as Counter");
                use.query.AppendLine(" from");
                use.query.AppendLine("    Pedido as pe");
                use.query.AppendLine("    inner join CuentaDeCompra as cc on cc.Correo = pe.Correo");
                use.query.AppendLine(" where");
                use.query.AppendLine("    pe.OrdenDeAmazon = @AmazonOrder");
                use.query.AppendLine("    and pe.Correo is not null");
                use.query.AppendLine(" group by");
                use.query.AppendLine("    pe.Correo,");
                use.query.AppendLine("    cc.EnviarCopiaExencion");

                execute.fillTable(ref use.query, ref tablita);

                if (tablita.Rows.Count == 0) { continue; }

                foreach (DataRow fila in tablita.Rows) {
                    if (fila["Counter"].ToString().Equals("0")) { continue; }
                    if (attachments.ToString().Trim().Equals("")) { continue; }

                    email.data.clearAll();
                    email.data.subject.Append("Tax exempt for ").Append(fila["Correo"]);
                    email.data.body.Append("<br/>Good day, attached proof of export.<br/><br/>E-mail account: ").Append(fila["Correo"]);
                    email.data.body.Append("<br/>Order numbers: ").Append(row["AmazonOrder"]);
                    email.data.body.Append("<br/><br/><br/>Best Regards,<br/>Mario Porres<br/><br/>");
                    email.data.origin.Append(fila["Correo"].ToString());
                    email.data.attachments.Append(attachments);

                    if (email.data.isSimulation) {
                        email.data.destination.Append("tec.desarrollo15.gtd@gmail.com");
                    } else {
                        email.data.destination.Append("tax-exempt@amazon.com");
                    }

                    if (fila["EnviarCopiaExencion"].ToString().Equals("1")) {
                        if (email.data.isSimulation) {
                            email.data.cc.Append("tec.teamoperaciones.gtd@gmail.com");
                        } else {
                            email.data.cc.Append(execute.getParameter(4));
                        }
                    }

                    if (!email.sendEmail(ref Error)) {
                        slack.data.Clear();
                        slack.data.alertTitle.Append("Error");
                        slack.data.subject.Append("Error procedure: error while sending mail, amazon order: ").Append(row["AmazonOrder"]).Append("");
                        slack.data.message.Append(Error);
                        slack.data.warningColour = slack.data.selectColour.Red;
                        slack.send();
                        continue;
                    }

                    use.query.Clear();
                    use.query.AppendLine(" Update");
                    use.query.AppendLine("    OrdenDeCompra");
                    use.query.AppendLine(" set");
                    use.query.AppendLine("    CorreoDeExencion = getdate()");
                    use.query.AppendLine(" where");
                    use.query.Append("    OrdenDeAmazon = '").Append(row["AmazonOrder"]).Append("'");
                    execute.executeQuery(ref use.query, "OrdeDeCompra has been update successfully! Amazon order:" + row["AmazonOrder"].ToString());

                    CounterGeneral++;
                }
                attachments.Clear();
            }
        }

        private Boolean verifyPackage(String AmazonOrder, ref StringBuilder TaxList, ref StringBuilder GuidesList, ref DataTable tablita) {
            use.query.Clear();
            use.query.Append("declare @AmazonOrder varchar(30) = '%").Append(AmazonOrder).AppendLine("%'");
            use.query.AppendLine(" select");
            use.query.AppendLine("    iif(");
            use.query.AppendLine("        charindex('(', pe.CodigoDeRastreo) > 0 and charindex(')', pe.CodigoDeRastreo) > 0,");
            use.query.AppendLine("        substring(");
            use.query.AppendLine("            pe.CodigoDeRastreo,");
            use.query.AppendLine("            charindex('(', pe.CodigoDeRastreo) + 1,");
            use.query.AppendLine("            charindex(')', pe.CodigoDeRastreo) - charindex('(', pe.CodigoDeRastreo) - 1");
            use.query.AppendLine("        ),");
            use.query.AppendLine("        pe.CodigoDeRastreo");
            use.query.AppendLine("    ) as ShortTracking,");
            use.query.AppendLine("    isnull(sum(pe.Impuesto), 0) as Tax,");
            use.query.AppendLine("    count(distinct pa.CodigoPaquete) as PackageCounter,");
            use.query.AppendLine("    pa.GuiaAerea as AirGuide");
            use.query.AppendLine(" from");
            use.query.AppendLine("    Pedido as pe");
            use.query.AppendLine("    left join Paquete as pa");
            use.query.AppendLine("        on pa.CodigoDeRastreo like iif(");
            use.query.AppendLine("            charindex('(', pe.CodigoDeRastreo) > 0 and charindex(')', pe.CodigoDeRastreo) > 0,");
            use.query.AppendLine("            substring(");
            use.query.AppendLine("                pe.CodigoDeRastreo,");
            use.query.AppendLine("                charindex('(', pe.CodigoDeRastreo) + 1,");
            use.query.AppendLine("                charindex(')', pe.CodigoDeRastreo) - charindex('(', pe.CodigoDeRastreo) - 1");
            use.query.AppendLine("            ),");
            use.query.AppendLine("            pe.CodigoDeRastreo");
            use.query.AppendLine("        )");
            use.query.AppendLine(" where");
            use.query.AppendLine("    pe.OrdenDeAmazon like @AmazonOrder");
            use.query.AppendLine("    and pe.CodigoEstadoPedido != 4");
            use.query.AppendLine("    and pe.CodigoDeRastreo is not null");
            use.query.AppendLine(" group by");
            use.query.AppendLine("    pe.CodigoDeRastreo,");
            use.query.AppendLine("    pa.GuiaAerea");

            execute.fillTable(ref use.query, ref tablita);

            if (tablita.Rows.Count == 0) { return false; }

            TaxList.Clear();
            int PackageCounter = 0;
            String AirGuide;
            use.query.Clear();

            foreach (DataRow row in tablita.Rows) {
                if (!row["Tax"].ToString().Equals("0.00")) { TaxList.Append(row["Tax"].ToString()).Append(","); }
                if (!row["PackageCounter"].ToString().Equals("1")) { continue; }

                PackageCounter++;

                AirGuide = row["AirGuide"].ToString().Trim();
                if (!AirGuide.Equals("")) {
                    if (!GuidesList.ToString().Contains(AirGuide)) { GuidesList.Append(AirGuide).Append(","); }

                    continue;
                }

                AirGuide = getNewGuide();

                use.query.AppendLine(" update");
                use.query.AppendLine("    Paquete");
                use.query.AppendLine(" set");
                use.query.Append("    GuiaAerea = '").Append(AirGuide).Append("',");
                use.query.AppendLine("    Generado = 1");
                use.query.AppendLine(" where");
                use.query.Append("    CodigoDeRastreo like '%").Append(row["ShortTracking"].ToString()).Append("%'");

                if (!GuidesList.ToString().Contains(AirGuide)) { GuidesList.Append(AirGuide).Append(","); }
            }

            if (!use.query.ToString().Equals("")) { execute.executeQuery(ref use.query, "Package-AirGuide has been updated successfully! Amazon order:" + AmazonOrder); }

            use.query.Clear();
            use.query.Append(" declare @AmazonOrder varchar(30) = '%").Append(AmazonOrder).AppendLine("%'");
            use.query.AppendLine(" select");
            use.query.AppendLine("    count(distinct pe.CodigoDeRastreo) as TotalPedidos");
            use.query.AppendLine(" from");
            use.query.AppendLine("    Pedido as pe");
            use.query.AppendLine(" where");
            use.query.AppendLine("    pe.CodigoDeRastreo is not null");
            use.query.AppendLine("    and pe.CodigoEstadoPedido != 4");
            use.query.AppendLine("    and pe.OrdenDeAmazon like @AmazonOrder");

            if (execute.getNat(ref use.query, 0) != PackageCounter) { return false; }

            GuidesList.Replace(GuidesList.ToString(), GuidesList.ToString().Substring(0, GuidesList.Length - 1));
            TaxList.Replace(TaxList.ToString(), TaxList.ToString().Substring(0, TaxList.Length - 1));

            return true;
        }

        private void generateGuide(DataRow row, ref DataTable tablita) {
            if (row["ProductsReceived"].ToString().Equals("0")) { return; }

            use.query.Clear();
            use.query.Append(" declare @AmazonOrder varchar(30) = '%").Append(row["AmazonOrder"]).AppendLine("%'");
            use.query.AppendLine(" select");
            use.query.AppendLine("    count(distinct pa.CodigoPaquete) as Packages,");
            use.query.AppendLine("    sum(iif(pe.CodigoPaquete is null, 1, 0)) as OrdersNoPackages,");
            use.query.AppendLine("    count(distinct pe.CodigoPedido) as Orders,");
            use.query.AppendLine("    iif(");
            use.query.AppendLine("        charindex('(', pe.CodigoDeRastreo) > 0 and charindex(')', pe.CodigoDeRastreo) > 0,");
            use.query.AppendLine("        substring(");
            use.query.AppendLine("            pe.CodigoDeRastreo,");
            use.query.AppendLine("            charindex('(', pe.CodigoDeRastreo) + 1,");
            use.query.AppendLine("            charindex(')', pe.CodigoDeRastreo) - charindex('(', pe.CodigoDeRastreo) - 1");
            use.query.AppendLine("        ),");
            use.query.AppendLine("        pe.CodigoDeRastreo");
            use.query.AppendLine("    )  as ShortTracking");
            use.query.AppendLine(" from");
            use.query.AppendLine("    Pedido as pe");
            use.query.AppendLine("    left join Paquete as pa");
            use.query.AppendLine("        on pa.CodigoDeRastreo like '%' + iif(");
            use.query.AppendLine("            charindex('(', pe.CodigoDeRastreo) > 0 and charindex(')', pe.CodigoDeRastreo) > 0,");
            use.query.AppendLine("            substring(");
            use.query.AppendLine("                pe.CodigoDeRastreo,");
            use.query.AppendLine("                charindex('(', pe.CodigoDeRastreo) + 1,");
            use.query.AppendLine("                charindex(')', pe.CodigoDeRastreo) - charindex('(', pe.CodigoDeRastreo) - 1");
            use.query.AppendLine("            ),");
            use.query.AppendLine("            pe.CodigoDeRastreo");
            use.query.AppendLine("        )+ '%'");
            use.query.AppendLine(" where");
            use.query.AppendLine("    pe.OrdenDeAmazon like @AmazonOrder");
            use.query.AppendLine(" group by");
            use.query.AppendLine("    pe.CodigoDeRastreo,");
            use.query.AppendLine("    pa.CodigoPaquete");

            execute.fillTable(ref use.query, ref tablita);

            if (tablita.Rows.Count == 0) { return; }

            foreach (DataRow fila in tablita.Rows) {
                if (fila["Packages"].ToString().Equals("0")) {
                    if (fila["Orders"].ToString().Equals("0")) { continue; }

                    insertPackage(fila["ShortTracking"].ToString(), row["AmazonOrder"].ToString());
                    updatePackage(fila["ShortTracking"].ToString());
                    continue;
                }

                if (fila["OrdersNoPackages"].ToString().Equals("0")) { continue; }

                use.query.Clear();
                use.query.Append(" declare @Tracking varchar(50) = '%").Append(fila["ShortTracking"].ToString()).AppendLine("%'");
                use.query.AppendLine(" update");
                use.query.AppendLine("    Pedido");
                use.query.AppendLine(" set");
                use.query.AppendLine("    CodigoPaquete = (Select max(CodigoPaquete) from Paquete where CodigoDeRastreo like @Tracking)");
                use.query.AppendLine(" where");
                use.query.AppendLine("    (CodigoDeRastreo = @AmazonOrder");
                use.query.AppendLine("    or CodigoDeRastreo like @Tracking)");
                use.query.AppendLine("    and CodigoPaquete is null");

                execute.executeQuery(ref use.query, "Order has been updated successfully! Amazon order: " + row["AmazonOrder"].ToString() + ", Tracking: " + fila["ShortTracking"].ToString());
            }
        }

        private void updatePackage(String Tracking) {
            if (Tracking.Trim().Equals("")) { return; }
            if (!validateTracking(Tracking).Equals("")) {
                Console.WriteLine("Tracking: " + Tracking + " doesn't exists or there is more than one record with this tracking.");
                return;
            }

            //Update order
            use.query.Clear();
            use.query.Append(" declare @Tracking varchar(30) = '").Append(Tracking).AppendLine("'");
            use.query.AppendLine(" update");
            use.query.AppendLine("    Pedido");
            use.query.AppendLine(" set");
            use.query.AppendLine("    CodigoPaquete = (");
            use.query.AppendLine("        select");
            use.query.AppendLine("            max(pa.CodigoPaquete) as CodigoPaquete");
            use.query.AppendLine("        from");
            use.query.AppendLine("            Paquete as pa");
            use.query.AppendLine("        where");
            use.query.AppendLine("            pa.CodigoDeRastreo like '%' + @Tracking + '%'");
            use.query.AppendLine("    )");
            use.query.AppendLine(" where");
            use.query.AppendLine("    CodigoDeRastreo = (");
            use.query.AppendLine("        select");
            use.query.AppendLine("            max(pe.CodigoDeRastreo) as CodigoDeRastreo");
            use.query.AppendLine("        from");
            use.query.AppendLine("            Pedido as pe");
            use.query.AppendLine("        where");
            use.query.AppendLine("            pe.CodigoDeRastreo = @Tracking");
            use.query.AppendLine("            or pe.CodigoDeRastreo like '%' + @Tracking + '%'");
            use.query.AppendLine("    )");
            //Update sales
            use.query.AppendLine(" update");
            use.query.AppendLine("    Venta");
            use.query.AppendLine(" Set");
            use.query.AppendLine("    CodigoEstadoEntrega = 3");
            use.query.AppendLine(" where");
            use.query.AppendLine("    CodigoEstadoEntrega = 2");
            use.query.AppendLine("    and CodigoVenta in (");
            use.query.AppendLine("        Select distinct");
            use.query.AppendLine("            vp.CodigoVenta");
            use.query.AppendLine("        from");
            use.query.AppendLine("            VentaPedido vp inner join Pedido p on vp.CodigoPedido = p.CodigoPedido");
            use.query.AppendLine("            inner join Venta v on vp.CodigoVenta = v.CodigoVenta");
            use.query.AppendLine("            inner join Paquete pa on p.CodigoPaquete = pa.CodigoPaquete");
            use.query.AppendLine("        where");
            use.query.AppendLine("            v.CodigoEstadoEntrega = 2");
            use.query.AppendLine("            And pa.GuiaAerea Is Not null");
            use.query.AppendLine("            And p.CodigoPaquete = (select max(CodigoPaquete) from Paquete where CodigoDeRastreo like '%' +@Tracking + '%')");
            use.query.AppendLine("    )");
            //Update order warehouse
            use.query.AppendLine(" update");
            use.query.AppendLine("    Pedido");
            use.query.AppendLine(" Set");
            use.query.AppendLine("    CodigoEstadoPedido = 2");
            use.query.AppendLine(" where");
            use.query.AppendLine("    CodigoEstadoPedido = 1");
            use.query.AppendLine("    and CodigoDeRastreo In (");
            use.query.AppendLine("        Select top(1)");
            use.query.AppendLine("            CodigoDeRastreo");
            use.query.AppendLine("        from");
            use.query.AppendLine("            Pedido");
            use.query.AppendLine("        where");
            use.query.AppendLine("            CodigoDeRastreo = @Tracking");
            use.query.AppendLine("            or CodigoDeRastreo like '%' + @Tracking + '%'");
            use.query.AppendLine("    )");

            execute.executeQuery(ref use.query, "Order and Sales have been updated successfully! Trackin: " + Tracking);
            email.sendMailTracking(ref Tracking, ref use.query, ref execute);
        }

        private String validateTracking(String Tracking) {
            if (Tracking.Length > 60) {
                return "Tracking is very long.";
            }

            use.query.Clear();
            use.query.Append(" declare @Tracking varchar(30) = '").Append(Tracking).AppendLine("'");
            use.query.AppendLine(" select");
            use.query.AppendLine("    count(1) as TrackingExists");
            use.query.AppendLine(" from");
            use.query.AppendLine("    Pedido as pe");
            use.query.AppendLine(" where");
            use.query.AppendLine("    pe.CodigoEstadoPedido != 4");
            use.query.AppendLine("    and (");
            use.query.AppendLine("        pe.CodigoDeRastreo = @Tracking");
            use.query.AppendLine("        or pe.CodigoDeRastreo like '%'  + @Tracking + '%'");
            use.query.AppendLine("    )");

            use.Nat = execute.getNat(ref use.query);

            if (use.Nat == 0) {
                return "Tracking doesn't exists.";
            }

            if (use.Nat > 1) {
                return "There is more than one record with this tracking.";
            }

            return "";
        }

        private void insertPackage(String Tracking, String AmazonOrder) {
            if (Tracking.Trim().Equals("")) { return; }

            use.query.Clear();
            use.query.Append(" declare @Rastreo varchar(50) = '%").Append(Tracking).AppendLine("%'");
            use.query.Append(" declare @AmazonOrder varchar(30) = '").Append(AmazonOrder).AppendLine("'");
            use.query.AppendLine(" select");
            use.query.AppendLine("    pe.CodigoDeRastreo,");
            use.query.AppendLine("    isnull(max(pe.Vendedor), '') as EnviadoPor,");
            use.query.AppendLine("    iif(");
            use.query.AppendLine("        trim(isnull(max(pe.TipoDeProducto), '')) != '',");
            use.query.AppendLine("        replace(max(pe.TipoDeProducto), ',', ' '),");
            use.query.AppendLine("        ''");
            use.query.AppendLine("    ) as Descripcion,");
            use.query.AppendLine("    sum(isnull(pr.Peso, 0)) as Peso,");
            use.query.AppendLine("    iif(pe.OrdenDeAmazon = @AmazonOrder, pr.CodigoEstablecimiento, 0) as CodigoEstablecimiento");
            use.query.AppendLine(" from");
            use.query.AppendLine("    Pedido as pe");
            use.query.AppendLine("    inner join Producto as pr on pr.CodigoAmazon = pe.CodigoAmazon");
            use.query.AppendLine(" where");
            use.query.AppendLine("    pe.CodigoDeRastreo like @Rastreo");
            use.query.AppendLine(" group by");
            use.query.AppendLine("    pe.CodigoDeRastreo,");
            use.query.AppendLine("    pe.OrdenDeAmazon,");
            use.query.AppendLine("    pr.CodigoEstablecimiento");
            DataTable table = new DataTable();
            execute.fillTable(ref use.query, ref table);

            if (table.Rows.Count == 0) { return; }
            use.query.Clear();
            foreach (DataRow row in use.tabBuffer.Rows) {
                use.query.AppendLine(" insert into");
                use.query.AppendLine("    Paquete (");
                use.query.AppendLine("        CodigoDeRastreo,");
                use.query.AppendLine("        GuiaAerea,");
                use.query.AppendLine("        EnviadoPor,");
                use.query.AppendLine("        Descripcion,");
                use.query.AppendLine("        Peso,");
                use.query.AppendLine("        Fecha,");
                use.query.AppendLine("        Generado,");
                use.query.AppendLine("        CodigoEstadoPaquete,");
                use.query.AppendLine("        CodigoEstablecimiento");
                use.query.AppendLine("    )values(");
                use.query.Append("        '").Append(Tracking).Append("',");
                use.query.Append("        '").Append(getNewGuide()).Append("',");
                use.query.Append("        '").Append(row["EnviadoPor"].ToString()).Append("',");
                use.query.Append("        '").Append(row["Descripcion"].ToString()).Append("',");
                use.query.Append("        ").Append(row["Peso"].ToString()).Append(",");
                use.query.Append("        getdate(),");
                use.query.Append("        1,");
                use.query.Append("        1,");
                use.query.Append("        ").Append(row["CodigoEstablecimiento"]);
                use.query.AppendLine(" )");
            }

            execute.executeQuery(ref use.query, "Package has been added successfully! Amazon order: " + Tracking);
        }

        private String getNewGuide() {
            use.query.Clear();
            use.query.AppendLine(" Declare @Guide int = (");
            use.query.AppendLine("    select");
            use.query.AppendLine("        max(");
            use.query.AppendLine("            isnull(");
            use.query.AppendLine("                TRY_CONVERT(int,");
            use.query.AppendLine("                    substring(");
            use.query.AppendLine("                        p.GuiaAerea,");
            use.query.AppendLine("                        2,");
            use.query.AppendLine("                        len(p.GuiaAerea)");
            use.query.AppendLine("                    )");
            use.query.AppendLine("                ), 0");
            use.query.AppendLine("            )");
            use.query.AppendLine("        ) as Guides");
            use.query.AppendLine("    from");
            use.query.AppendLine("        Paquete as p");
            use.query.AppendLine("    where");
            use.query.AppendLine("        p.GuiaAerea like 'G%'");
            use.query.AppendLine(" )");
            use.query.AppendLine(" update Parametro set Valor = @Guide + 1 where CodigoParametro = 51");
            use.query.AppendLine(" select @Guide + 1 as Result");

            return "G" + execute.getNat(ref use.query);
        }
    }
}

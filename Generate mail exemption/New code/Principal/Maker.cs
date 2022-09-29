using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace Principal{
    internal class AsynClss { }
    internal class Maker{
        private common.UseCommon use = new common.UseCommon();
        private connection.Execute execute = new connection.Execute();
        private DataTable tablita = new DataTable();

        private static Task<Principal> getAsyn() { 
        }

        public void makeAll() {
            Console.WriteLine("Inicializando...");
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
            use.query.AppendLine("                case when (");
            use.query.AppendLine("                    p.Cantidad > 0");
            use.query.AppendLine("                    and p.CodigoEstadoPedido in (1, 2, 6)");
            use.query.AppendLine("                ) then 1 else 0 end");
            use.query.AppendLine("            ), 0");
            use.query.AppendLine("        ) = 0");

            use.query.AppendLine(" select");
            use.query.AppendLine("    pe.OrdenDeAmazon as AmazonOrder,");
            use.query.AppendLine("    isnull(");
            use.query.AppendLine("        sum(");
            use.query.AppendLine("            case when (");
            use.query.AppendLine("                pe.CodigoEstadoPedido = 3");
            use.query.AppendLine("            ) then 1 else 0 end");
            use.query.AppendLine("        ), 0");
            use.query.AppendLine("    ) as ProductsReceived,");
            use.query.AppendLine("    case");
            use.query.AppendLine("        when (");
            use.query.AppendLine("            charindex('(', pe.CodigoDeRastreo) > 0");
            use.query.AppendLine("            and charindex(')', pe.CodigoDeRastreo) > 0");
            use.query.AppendLine("        ) then (");
            use.query.AppendLine("            substring(");
            use.query.AppendLine("                pe.CodigoDeRastreo,");
            use.query.AppendLine("                charindex('(', pe.CodigoDeRastreo) + 1,");
            use.query.AppendLine("                charindex(')', pe.CodigoDeRastreo) - charindex('(', pe.CodigoDeRastreo) - 1");
            use.query.AppendLine("            )");
            use.query.AppendLine("        ) else pe.CodigoDeRastreo");
            use.query.AppendLine("    end as ShortTracking");
            use.query.AppendLine(" from");
            use.query.AppendLine("    Pedido as pe");
            use.query.AppendLine("    inner join CuentaDeCompra as cc on cc.Correo = pe.Correo");
            use.query.AppendLine("    inner join @Temp as t on t.AmazonOrder = pe.OrdenDeAmazon");
            use.query.AppendLine(" group by");
            use.query.AppendLine("    pe.OrdenDeAmazon,");
            use.query.AppendLine("    pe.CodigoDeRastreo");
            use.query.AppendLine(" having");
            use.query.AppendLine("    (");
            use.query.AppendLine("        isnull(SUM(case when (pe.CodigoDeRastreo is not null) then pe.Impuesto else 0 end),0) > 0");
            use.query.AppendLine("        or isnull(sum(case when (cc.SiempreEnviarExencion = 1) then 1 else 0 end), 0) > 0");
            use.query.AppendLine("    )");
            use.query.AppendLine("    and isnull(sum(case when (cc.NuncaEnviarExencion = 1) then 1 else 0 end), 0) = 0");
            use.query.AppendLine("    and isnull(sum(case when (pe.Cantidad > 0 and pe.CodigoEstadoPedido != 4) then 1 else 0 end), 0) > 0");

            execute.fillTable(ref use.query, ref use.tabBuffer);

            Console.WriteLine("Rows: " + use.tabBuffer.Rows.Count.ToString());
            Console.WriteLine("Columns: " + use.tabBuffer.Rows.Count.ToString());

            foreach (DataRow row in use.tabBuffer.Rows) {
                generateGuide(row);
            }
        }

        private void generateGuide(DataRow row) {
            if (row["ProductsReceived"].ToString().Equals("0")) { return; }

            use.query.Clear();
            use.query.Append(" declare @AmazonOrder varchar(50) = '").Append(row["ShortTracking"].ToString()).Append("'");
            use.query.AppendLine("    select");
            use.query.AppendLine("        count(1) as Counter,");
            use.query.AppendLine("        'Paquete' as Result");
            use.query.AppendLine("    from");
            use.query.AppendLine("        Paquete");
            use.query.AppendLine("    where");
            use.query.AppendLine("        CodigoDeRastreo like '%' + @AmazonOrder + '%'");
            use.query.AppendLine(" union");
            use.query.AppendLine("    select");
            use.query.AppendLine("        Count(1) as Counter,");
            use.query.AppendLine("        'Pedido' as Result");
            use.query.AppendLine("    from");
            use.query.AppendLine("        Pedido");
            use.query.AppendLine("    where");
            use.query.AppendLine("        CodigoDeRastreo like '%' + @AmazonOrder + '%'");
            tablita.Clear();
            execute.fillTable(ref use.query, ref tablita); 

            if (!tablita.Rows[0]["Counter"].ToString().Equals("0")) {
                if (tablita.Rows[1]["Counter"].ToString().Equals("0")) { return; }

                use.query.Clear();
                use.query.Append(" declare @AmazonOrder varchar(50) = '").Append(row["ShortTracking"].ToString()).Append("'");
                use.query.AppendLine(" update");
                use.query.AppendLine("    Pedido");
                use.query.AppendLine(" set");
                use.query.AppendLine("    CodigoPaquete = (Select top 1 CodigoPaquete from Paquete where CodigoDeRastreo like '%' + @AmazonOrder + '%')");
                use.query.AppendLine(" where");
                use.query.AppendLine("    CodigoDeRastreo = @AmazonOrder");
                use.query.AppendLine("    or CodigoDeRastreo like '%' + @AmazonOrder + '%'");

                Console.WriteLine("Updated order with AmazonrOrder: " + row["ShortTracking"].ToString()) ;
                //execute.executeQuery(ref use.query);
                return;
            }

            if (tablita.Rows[1]["Counter"].ToString().Equals("0")) { return; }

            //Insertarl paquete(codigoPaquete, precio, descripcion...)
            Console.WriteLine("Insert package with AmazonOrder " + row["ShortTracking"].ToString());
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

 declare @Temp table(AmazonOrder varchar(50))
 insert into @Temp
    select
        oc.OrdenDeAmazon
    from
        OrdenDeCompra as oc
        inner join Pedido as p on p.OrdenDeAmazon = oc.OrdenDeAmazon
    where
        oc.CorreoDeExencion is null
    group by
        oc.OrdenDeAmazon
    having
        isnull(
            sum(
                iif(p.Cantidad > 0 and p.CodigoEstadoPedido in (1, 2, 6), 1, 0)
            ), 0
        ) = 0

 select
    pe.OrdenDeAmazon as AmazonOrder,
    isnull(
        sum(
            iif(pe.CodigoEstadoPedido = 3, 1, 0)
        ), 0
    ) as ProductsReceived
 from
    Pedido as pe
    inner join CuentaDeCompra as cc on cc.Correo = pe.Correo
    inner join @Temp as t on t.AmazonOrder = pe.OrdenDeAmazon
 group by
    pe.OrdenDeAmazon
 having
    (
        isnull(SUM(iif(pe.CodigoDeRastreo is not null, pe.Impuesto, 0)),0) > 0
        or isnull(sum(iif(cc.SiempreEnviarExencion = 1, 1, 0)), 0) > 0
    )
    and isnull(sum(iif(cc.NuncaEnviarExencion = 1, 1, 0)), 0) = 0
    and isnull(sum(iif(pe.Cantidad > 0 and pe.CodigoEstadoPedido != 4, 1, 0)), 0) > 0
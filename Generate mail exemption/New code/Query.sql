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
                case when (
                    p.Cantidad > 0
                    and p.CodigoEstadoPedido in (1, 2, 6)
                ) then 1 else 0 end
            ), 0
        ) = 0
            
select
    pe.OrdenDeAmazon as AmazonOrder,
    isnull(
        sum(
            case when (
                pe.CodigoEstadoPedido = 3
            ) then 1 else 0 end
        ), 0
    ) as ProductsReceived,
    case 
        when (
            charindex('(', pe.CodigoDeRastreo) > 0
            and charindex(')', pe.CodigoDeRastreo) > 0
        ) then (
            substring(
                pe.CodigoDeRastreo,
                charindex('(', pe.CodigoDeRastreo) + 1,
                charindex(')', pe.CodigoDeRastreo) - charindex('(', pe.CodigoDeRastreo) - 1
            )
        ) else pe.CodigoDeRastreo
    end as ShortTracking
from
    Pedido as pe
    inner join CuentaDeCompra as cc on cc.Correo = pe.Correo
    inner join @Temp as t on t.AmazonOrder = pe.OrdenDeAmazon
group by
    pe.OrdenDeAmazon,
    pe.CodigoDeRastreo
having
    (
        isnull(SUM(case when (pe.CodigoDeRastreo is not null) then pe.Impuesto else 0 end),0) > 0
        or isnull(sum(case when (cc.SiempreEnviarExencion = 1) then 1 else 0 end), 0) > 0
    )
    and isnull(sum(case when (cc.NuncaEnviarExencion = 1) then 1 else 0 end), 0) = 0
    and isnull(sum(case when (pe.Cantidad > 0 and pe.CodigoEstadoPedido != 4) then 1 else 0 end), 0) > 0

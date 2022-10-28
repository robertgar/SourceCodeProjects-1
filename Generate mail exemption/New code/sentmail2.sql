declare @SalesCode int = (
    select
        max(vp.CodigoVenta)
    from
        VentaPedido as vp
        inner join Pedido as pe on pe.CodigoPedido = vp.CodigoPedido
        inner join Paquete as pa on pa.CodigoPaquete = pe.CodigoPaquete
    where
        pa.CodigoDeRastreo like '%TBA164758546704%'
)

select
    ee.CodigoEstadoEntrega,
    ee.Nombre,
    ee.Etapa,
    dt.Datita
from
    EstadoEntrega as ee
    inner join (
        select
            substring(sp.Value, 1, charindex(':', sp.Value) - 1) as DeliveryStatusCode,
            substring(sp.Value, charindex(':', sp.Value) + 1, len(sp.Value)) as Datita
        from
            string_split(
                (
                    select
                        iif(pe.FechaDeOrden is not null, ',1:' + convert(varchar, pe.FechaDeOrden, 103), '') +
                        iif(pe.FechaDeEnvio is not null, ',2:' + convert(varchar, pe.FechaDeEnvio, 103), '') +
                        iif(pa.Fecha is not null, ',3:' + convert(varchar, pa.Fecha, 103), '') +
                        iif(
                            iif(pa.CodigoPaquete = pe.CodigoPaquete and pa.FechaRecibida is not null, pa.FechaRecibida, v.FechaConfirmacion) is not null, 
                                ',4:' + convert(varchar, iif(pa.CodigoPaquete = pe.CodigoPaquete and pa.FechaRecibida is not null, pa.FechaRecibida, v.FechaConfirmacion), 103)
                                , ''
                        ) +
                        iif(f.Fecha is not null, ',5:'+convert(varchar, f.Fecha, 103), '') +
                        iif(v.FechaEntrega is not null, ',6:' + convert(varchar, v.FechaEntrega, 103), '') as Result
                    from
                        Venta as v
                        inner join VentaPedido as vp on vp.CodigoVenta = v.CodigoVenta
                        inner join Pedido as pe on pe.CodigoPedido = vp.CodigoPedido
                        left join Paquete as pa on pa.CodigoPaquete = pe.CodigoPaquete
                        inner join Factura as f on f.CodigoFactura = v.CodigoFactura
                    where
                        v.CodigoVenta = @SalesCode
                ), ','
            ) as sp
        where
            trim(sp.Value) != ''
    ) as dt on dt.DeliveryStatusCode = ee.CodigoEstadoEntrega
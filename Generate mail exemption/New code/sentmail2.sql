declare @SalesCode int = 222320

select
    substring(sp.[value], 1, charindex(':', sp.[value]) - 1)
from
    string_split(
        (
            select
                iif(pe.FechaDeOrden is not null, ',1:' + convert(varchar, pe.FechaDeOrden, 1), '') +
                iif(pe.FechaDeEnvio is not null, ',2:' + convert(varchar, pe.FechaDeEnvio, 1), '') +
                iif(pa.Fecha is not null, ',3:' + convert(varchar, pa.Fecha, 1), '') +
                iif(
                    iif(pa.CodigoPaquete = pe.CodigoPaquete and pa.FechaRecibida is not null, pa.FechaRecibida, v.FechaConfirmacion) is not null, 
                        ',4:' + convert(varchar, iif(pa.CodigoPaquete = pe.CodigoPaquete and pa.FechaRecibida is not null, pa.FechaRecibida, v.FechaConfirmacion), 1)
                        , ''
                ) +
                iif(f.Fecha is not null, ',5:'+convert(varchar, f.Fecha, 1), '') +
                iif(v.FechaEntrega is not null, ',6:' + convert(varchar, v.FechaEntrega, 1), '') as Result
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
    sp.[value] != ''

return
select
    CodigoEstadoEntrega,
    nombre,
    Etapa, 
    ISNULL(
        CASE 
            WHEN CodigoEstadoEntrega = 1
                THEN (
                    SELECT 
                        CONVERT(varchar(10), CONVERT(date, fechadeorden, 106), 103)
                    FROM 
                        Pedido 
                    WHERE 
                        CodigoPedido = (
                            SELECT 
                                CodigoPedido 
                            FROM 
                                VentaPedido
                            WHERE 
                                CodigoVenta = @SalesCode
                        )
                ) 
            WHEN CodigoEstadoEntrega = 2 THEN (
                SELECT
                    CONVERT(varchar(10), CONVERT(date, FechaDeEnvio, 106), 103) 
                FROM 
                    Pedido
                WHERE 
                    CodigoPedido = (
                        SELECT 
                            CodigoPedido 
                        FROM 
                            VentaPedido 
                        WHERE 
                            CodigoVenta = @SalesCode
                    )
            ) 
            WHEN CodigoEstadoEntrega = 3 THEN (
                SELECT 
                    CONVERT(varchar(10), CONVERT(date, Fecha, 106), 103)
                FROM 
                    Paquete 
                WHERE 
                    CodigoPaquete = (
                        SELECT 
                            p.CodigoPaquete 
                        FROM 
                            Pedido p
                            INNER JOIN VentaPedido vp ON p.CodigoPedido = vp.CodigoPedido 
                        WHERE 
                            vp.CodigoVenta = @SalesCode
                    )
            )
            WHEN CodigoEstadoEntrega = 4 THEN (
                case when exists (
                    SELECT 
                        FechaRecibida 
                    FROM 
                        Paquete 
                    WHERE 
                        CodigoPaquete = (
                            SELECT 
                                p.CodigoPaquete
                            FROM 
                                Pedido p 
                                INNER JOIN VentaPedido vp ON p.CodigoPedido = vp.CodigoPedido 
                            WHERE 
                                vp.CodigoVenta = @SalesCode
                        )
                ) then (
                    SELECT 
                        CONVERT(varchar(10), CONVERT(date, FechaRecibida, 106), 103) 
                    FROM 
                        Paquete 
                    WHERE 
                        CodigoPaquete = (
                            SELECT 
                                p.CodigoPaquete 
                            FROM 
                                Pedido p
                                INNER JOIN VentaPedido vp ON p.CodigoPedido = vp.CodigoPedido 
                            WHERE 
                                vp.CodigoVenta = @SalesCode
                        )
                ) else (
                    select 
                        CONVERT(varchar(10), CONVERT(date, fechaconfirmacion, 106), 103) 
                    from 
                        venta 
                    where 
                        codigoventa = @SalesCode
                ) end
            )
            WHEN CodigoEstadoEntrega = 5 THEN (
                SELECT 
                    CONVERT(varchar(10), CONVERT(date, Fecha, 106), 103)
                FROM 
                    Factura 
                WHERE 
                    CodigoFactura = (
                        SELECT 
                            CodigoFactura 
                        FROM 
                            Venta 
                        WHERE 
                            CodigoVenta = @SalesCode
                    )
            )
            WHEN CodigoEstadoEntrega = 6 THEN (
                SELECT 
                    CONVERT(varchar(10), CONVERT(date, FechaEntrega, 106), 103)
                FROM 
                    Venta 
                WHERE 
                    CodigoVenta = @SalesCode
            ) 
        END, ''
    ) AS fecha 
FROM 
    EstadoEntrega 
WHERE 
    CodigoEstadoEntrega <= (select cast(codigoestadoentrega as int) from Venta where CodigoVenta = @SalesCode)
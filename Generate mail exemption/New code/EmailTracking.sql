declare @Tracking varchar(30) =  '%TBA164758546704%'

select distinct
    v.CodigoVenta,
    v.CodigoProducto,
    convert(varchar, v.FechaConfirmacion, 103) as FechaConfirmacion,
    concat(isnull(v.Cantidad, 0), ' ', trim(v.NombreProducto) ) as NombreProducto,
    v.NombreCliente,
    v.CorreoCliente,
    v.Telefonos,
    isnull(v.EsPedido, 0) as EsPedido,
    concat(v.DireccionDeEntrega, ' Departamento: ', d.Nombre, ' Municipio:', m.Nombre) as DireccionDeEntrega,
    v.CodigoFormaDePago,
    fp.Nombre as FormaDePago,
    v.Monto,
    v.Envio,
    isnull(v.Cuotas, 1) as Cuotas,
    case v.CodigoFormaDePago
        when 1 then isnull(v.ServicioPagoEfectivo, 0)
        when 3 then isnull(v.MontoCuota, 0) - isnull(v.Monto, 0)
        else 0
    end as ServicioPago,
    (
        isnull(v.Monto, 0) + isnull(v.Envio, 0) +(
            case v.CodigoFormaDePago
                when 1 then isnull(v.ServicioPagoEfectivo, 0)
                when 3 then isnull(v.MontoCuota, 0) - isnull(v.Monto, 0)
                else 0
            end
        )
    ) as Total,
    v.Factura,
    pr.Foto,
    v.NombreRecibido,
    convert(varchar, v.FechaEntrega, 103) as FechaEntrega,
    ee.Nombre as EmpresaDeEntrega,
    trim(
        iif(
            charindex('NIT:', v.Factura) = 0,
            iif(
                charindex('DIR:', v.Factura) = 0,
                v.Factura,
                substring(v.Factura, 0, charindex('DIR:', v.Factura))
            ),
            substring(v.Factura, 0, charindex('NIT:', v.Factura))
        )
    ) as NombreFactura,
    trim(
        iif(
            charindex('NIT:', v.Factura) = 0,
            '',
            substring(
                v.Factura, 
                charindex('NIT:', v.Factura) + 4,
                iif(
                    charindex('DIR:', v.Factura) = 0,
                    len(v.Factura) - charindex('NIT:', v.Factura) - 4,
                    charindex('DIR:', v.Factura) - charindex('NIT:', v.Factura) - 4
                )
            )
        )
    ) as NitFactura,
    trim(
        iif(
            charindex('DIR:', v.Factura) = 0,
            '',
            substring(v.Factura, charindex('DIR:', v.Factura) + 4, len(v.Factura))
        )
    ) as DireccionFactura,
    concat('https://guatemaladigital.com/Producto/', pr.CodigoProducto) as Link,
    case 
        when v.EsPedido IS Not null then 6 
        when v.EsPedido IS NULL and v.CodigoDeRastreo IS not null then 3 
        when v.EsPedido IS NULL and v.CodigoDeRastreo IS null then 2 
    end as Steps
from
    Venta as v
    inner join VentaPedido as vp on vp.CodigoVenta = v.CodigoVenta
    inner join Pedido as pe on pe.CodigoPedido = vp.CodigoPedido
    inner join Paquete as pa on pa.CodigoPaquete = pe.CodigoPaquete
    inner join FormaDePago as fp on fp.CodigoFormaDePago = v.CodigoFormaDePago
    inner join Producto as pr on pr.CodigoProducto = v.CodigoProducto
    left join Departamento as d on d.CodigoDepartamento = v.CodigoDepartamento
    left join Municipio as m on m.CodigoMunicipio = v.CodigoMunicipio
    left join EmpresaDeEntrega as ee on ee.CodigoEmpresaDeEntrega = v.CodigoEmpresaDeEntrega
where
    v.CodigoEstadoDeVenta = 1
    and v.CodigoEstadoEntrega = 2
    and pa.GuiaAerea is not null
    and pa.CodigoDeRastreo like @Tracking
order by
    v.CodigoVenta desc
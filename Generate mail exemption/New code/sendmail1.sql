Declare @Sale int = (
    select
        max(pa.CodigoPaquete)
    from
        Paquete as pa
    where
        pa.CodigoDeRastreo like '%1Z0638E81316951892%'
)
select
    v.CodigoVenta, 
    v.CodigoProducto,
    v.FechaConfirmacion,
    convert(varchar, v.Cantidad) + ' ' + v.NombreProducto as NombreProducto,
    v.NombreCliente,
    v.CorreoCliente,
    v.Telefonos,
    isnull(convert(int, v.EsPedido), 0) as EsPedido,
    v.DireccionDeEntrega + ' Departamento: ' + d.Nombre + ' Municipio: ' + m.Nombre as DireccionDeEntrega,
    v.CodigoFormaDePago,
    fp.Nombre as FormaDePago,
    v.Monto,
    v.Envio,
    isnull(v.Cuotas, 1) as Cuotas,
    (        
        case v.CodigoFormaDePago
            when 1 then v.ServicioPagoEfectivo
            when 3 then v.MontoCuota - v.Monto
        else
            0.00
        end
    ) as ServicioPago,
    (
        v.Monto + v.Envio + (
            case v.CodigoFormaDePago
                when 1 then v.ServicioPagoEfectivo
                when 3 then v.MontoCuota - v.Monto
            else
                0.00
            end
        )
    ) as Total,
    v.Factura,
    p.Foto,
    v.NombreRecibido,
    v.FechaEntrega,
    ee.Nombre as EmpresaDeEntrega
from
    Venta as v
    inner join FormaDePago as fp on fp.CodigoFormaDePago = v.CodigoFormaDePago
    inner join Producto as p on p.CodigoProducto = v.CodigoProducto
    left join EmpresaDeEntrega as ee on ee.CodigoEmpresaDeEntrega = v.CodigoEmpresaDeEntrega
    left join Departamento as d on d.CodigoDepartamento = v.CodigoDepartamento
    left join Municipio as m on m.CodigoMunicipio = v.CodigoMunicipio
where
    v.CodigoVenta = @Sale
    and v.CodigoEstadoDeVenta = 1
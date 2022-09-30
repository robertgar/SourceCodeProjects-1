declare @Rastreo varchar(50) = '%1Z001A271302826046%'
declare @AmazonOrder varchar(30) = '111-4701190-5089821'

Declare @Guide int = (
    select
        max(
            isnull(
                TRY_CONVERT(int,
                    substring(
                        p.GuiaAerea,
                        2,
                        len(p.GuiaAerea)
                    )
                ), 0
            )
        ) as Guides
    from
        Paquete as p
    where
        p.GuiaAerea like 'G%'
) + 1

select
    pe.CodigoDeRastreo,
    @Guide,
    isnull(max(pe.Vendedor), '') as EnviadoPor,
    iif(
        trim(isnull(max(pe.TipoDeProducto), '')) != '',
        replace(max(pe.TipoDeProducto), ',', ' '),
        ''
    ) as Descripcion,
    sum(isnull(pr.Peso, 0)) as Peso,
    iif(pe.OrdenDeAmazon = @AmazonOrder, pr.CodigoEstablecimiento, 0) as CodigoEstablecimiento,
    max(pa.CodigoPaquete) as CodigoPaquete
from
    Pedido as pe
    inner join Producto as pr on pr.CodigoAmazon = pe.CodigoAmazon
    inner join Paquete as pa on pa.CodigoDeRastreo like @Rastreo
where
    pe.CodigoDeRastreo like @Rastreo
group by
    pe.CodigoDeRastreo,
    pe.OrdenDeAmazon,
    pr.CodigoEstablecimiento


select 
    @Rastreo, 
    @Guide,
    (
        select top 1 
            isnull(Vendedor,'') 
        from 
            Pedido 
        where 
            CodigoDeRastreo like @Rastreo
    ),(
        select top 1
            replace(p.TipoDeProducto, ',', ' ')
        from
            Pedido as p
        where
            p.CodigoDeRastreo like @Rastreo
            and trim(isnull(p.TipoDeProducto, '')) != ''
    ),(
        select 
            sum(isnull(pr.Peso,0)) as R
        from 
            Producto as pr
            inner join Pedido as pe
                on pe.CodigoAmazon = pr.CodigoAmazon
        where
            pe.CodigoDeRastreo like @Rastreo
    ),(
        select top(1) 
            pr.CodigoEstablecimiento
        from 
            Pedido pe
            inner join Producto pr on pr.CodigoAmazon = pe.CodigoAmazon
        where 
            pe.OrdenDeAmazon = @AmazonOrder
            and pe.CodigoDeRastreo is not null
    ),(
        select 
            CodigoPaquete 
        from 
            Paquete 
        where 
             CodigoDeRastreo like @Rastreo
    )

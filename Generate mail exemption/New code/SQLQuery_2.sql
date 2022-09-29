declare @Rastreo varchar(50) = '%1Z001A271302826046%'
declare @AmazonOrder varchar(30) = '111-4701190-5089821'

declare @EstablecimientoCode int  = (
select top(1) 
    pr.CodigoEstablecimiento
from 
    Pedido pe
    inner join Producto pr on pr.CodigoAmazon = pe.CodigoAmazon
where 
    pe.OrdenDeAmazon = @AmazonOrder
    and CodigoDeRastreo is not null
)

declare @Peso decimal(18, 2) = (
    select 
        sum(isnull(pr.Peso,0)) as R
    from 
        Producto as pr
        inner join Pedido as pe
            on pe.CodigoAmazon = pr.CodigoAmazon
    where
        pe.CodigoDeRastreo like @Rastreo
)

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
declare @Saler varchar(50) = (select top 1 isnull(Vendedor,'') from Pedido where CodigoDeRastreo like @Rastreo)

DECLARE @Description VARCHAR(510) = (
    select top 1
        replace(p.TipoDeProducto, ',', ' ')
    from
        Pedido as p
    where
        p.CodigoDeRastreo like @Rastreo
        and trim(isnull(p.TipoDeProducto, '')) != ''
)

Select 
    @Rastreo as CodigoDeRastreo,
    @Guide as Guia,
    @Saler as EnviadoPor

public class EnvioData
{
	public string Total { get; set; }
	public string NumeroTarjeta { get; set; }
	public string Mes { get; set; }
	public string Anio { get; set; }
	public string Cvv { get; set; }
	public string CodigoFactura { get; set; }

	public EnvioData(string Total1, string NumeroTarjeta1, string Mes1, string Anio1, string Cvv1, string CodigoFactura1)
	{
		Total = Total1;
		NumeroTarjeta = NumeroTarjeta1;
		Mes = Mes1;
		Anio = Anio1;
		Cvv = Cvv1;
		CodigoFactura = CodigoFactura1;
	}
}

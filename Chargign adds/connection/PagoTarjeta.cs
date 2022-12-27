
public class PagoTarjeta
{
    public int Status { get; set; }
    public int Code { get; set; }
    public ResponseCredomatic Response { get; set; }
}

public class ResponseCredomatic
{
    public string authcode { get; set; }
    public string transactionid { get; set; }
    public string respuestatext { get; set; }
    public string Message { get; set; }
}


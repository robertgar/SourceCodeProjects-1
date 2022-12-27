#nullable enable

using Newtonsoft.Json;
using RestSharp;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Threading.Tasks;

public static class CobroPeti
{
   public static async Task<PagoTarjeta?> AutoCredomaticAsync(EnvioData data)
    {
        var client = WebApi.GetApi().Client;
        var rq = new RestRequest("api/CobroAnuncios");
        string serializado = JsonConvert.SerializeObject(data, Newtonsoft.Json.Formatting.Indented);
        rq.AddHeader("Content-Type", "application/json");
        rq.AddParameter("application/json", serializado, ParameterType.RequestBody);
        var r = await client.ExecutePostAsync(rq);
        return r.GetResultInternalApi<PagoTarjeta>();
    }
    public static PagoTarjeta? AutoCredomatic(EnvioData data)
    {
        return Task.Run(async () => await AutoCredomaticAsync(data)).Result;
    }

}

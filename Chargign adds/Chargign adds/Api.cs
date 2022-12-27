#nullable enable

using Newtonsoft.Json;
using RestSharp;
using RestSharp.Serializers.NewtonsoftJson;
using System;
using System.Configuration;
using System.Net;
using System.Threading.Tasks;

public class BackendApi
{
    private static BackendApi API = null!;

    public static BackendApi GetApi() => API ?? (API = new BackendApi());

    public RestClient Client { get; private set; }

    public BackendApi()
    {
        var options = new RestClientOptions(https://guatemaladigital.org:86/);
        Client = new RestClient(options);
        Client.UseNewtonsoftJson();
    }

}

public class WebApi
{
    private static WebApi API = null!;

    public static WebApi GetApi() => API ?? (API = new WebApi());

    public RestClient Client { get; private set; }

    public WebApi()
    {
        var options = new RestClientOptions(https://guatemaladigital.org:86/);
        Client = new RestClient(options);
        Client.UseNewtonsoftJson();
    }

}

public class ApiResponse<T>
{
    public T? Data { get; set; }
    public bool OK { get; set; }
    public ApiError? Error { get; set; }
    public string Correlation { get; set; } = null!;
}

public class ApiError
{
    public string Title { get; }
    public string Message { get; }

    public ApiError(string title, string message)
    {
        Title = title;
        Message = message;
    }
}

public static class ApiExtensions
{

    public async static Task<bool> GetResult(this RestResponse request)
    {
        var result = await request.GetResultInternal<object>();

        if (!string.IsNullOrEmpty(request.ErrorMessage))
        {
            ShowNotification("Network Error", request.ErrorMessage);
            return false;
        }
        if (result.OK)
        {
            return true;
        }
        else
        {
            ShowNotification(result.Error?.Title ?? "", result.Error?.Message ?? "");
            return false;
        }
    }
    public async static Task<bool> GetResultNoError(this RestResponse request)
    {
        var result = await request.GetResultInternal<object>();
        if (!string.IsNullOrEmpty(request.ErrorMessage))
            return false;
        return result.OK;
    }

    public async static Task<T?> GetResult<T>(this RestResponse request)
    {
        var result = await request.GetResultInternal<T>();

        if (!string.IsNullOrEmpty(request.ErrorMessage))
        {
            ShowNotification("Network Error", request.ErrorMessage);
            return default(T);
        }
        if (result.OK)
        {
            return result.Data;
        }
        else
        {
            ShowNotification(result.Error?.Title ?? "", result.Error?.Message ?? "");
            return default(T);
        }
    }

    public async static Task<T?> GetResultNoError<T>(this RestResponse request)
    {
        var result = await request.GetResultInternal<T>();
        if (!string.IsNullOrEmpty(request.ErrorMessage))
            return default(T);
        if (result.OK)
            return result.Data;
        else
            return default(T);
    }

    public static T? GetResultInternalApi<T>(this RestResponse restResponse)
    {
        return JsonConvert.DeserializeObject<T>(restResponse.Content!);
    }
    public static async Task<ApiResponse<T>> GetResultInternal<T>(this RestResponse restResponse)
    {
        var apiResponse = new ApiResponse<T>();
        var data = await Task.Run(() =>
        {
            ApiResponse<T>? value = default;
            Exception? ex = null;
            try
            {
                if (restResponse.Content != null)
                    value = JsonConvert.DeserializeObject<ApiResponse<T>>(restResponse.Content);
            }
            catch (Exception e)
            {
                ex = e;
            }
            return new Tuple<ApiResponse<T>?, Exception?>(value, ex);
        });
        if (data.Item1 != null)
            return data.Item1;
        else
            apiResponse.Error = new ApiError("Error while decoding response", data.Item2?.Message ?? "");

        return apiResponse;
    }

    private static void ShowNotification(string title, string? content)
    {

    }
}
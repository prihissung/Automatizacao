using System.Net.Http;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;

namespace AutomacaoVestib.App.Automacao;

public class HttpIntegrationService
{
	private readonly HttpClient _httpClient;

	public HttpIntegrationService(HttpClient httpClient)
	{
		_httpClient = httpClient;
	}

	public async Task<HttpResponseMessage> SendAsync(IntegrationConfig config, string codConc, GeneratedValues values, CancellationToken ct)
	{
		string payload = config.PayloadTemplate
			.Replace("{CodConc}", codConc)
			.Replace("{Type6Value}", values.Type6Value)
			.Replace("{Type13Value}", values.Type13Value);

		var req = new HttpRequestMessage(new HttpMethod(config.Method), config.Url)
		{
			Content = new StringContent(payload, Encoding.UTF8, "application/json")
		};
		foreach (var kv in config.Headers)
		{
			req.Headers.TryAddWithoutValidation(kv.Key, kv.Value);
		}
		return await _httpClient.SendAsync(req, ct);
	}
}
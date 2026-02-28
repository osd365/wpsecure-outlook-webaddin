
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using Microsoft.IdentityModel.Protocols;
using Microsoft.IdentityModel.Protocols.OpenIdConnect;
using Microsoft.IdentityModel.Tokens;
using FunctionsHttpRequestData = Microsoft.Azure.Functions.Worker.Http.HttpRequestData;

namespace WPSecure.Api.Signatures
{
    public sealed class BootstrapRequest
    {
        [JsonPropertyName("id_token")] public string? IdToken { get; set; }
    }

    public sealed class BootstrapFiles
    {
        [JsonPropertyName("newHtml")] public string? NewHtml { get; set; }
        [JsonPropertyName("replyHtml")] public string? ReplyHtml { get; set; }
        [JsonPropertyName("newText")] public string? NewText { get; set; }
        [JsonPropertyName("replyText")] public string? ReplyText { get; set; }
        [JsonPropertyName("apptNewHtml")] public string? ApptNewHtml { get; set; }
        [JsonPropertyName("apptNewText")] public string? ApptNewText { get; set; }
    }

    public sealed class BootstrapResult
    {
        [JsonPropertyName("success")] public bool Success { get; set; }
        [JsonPropertyName("error")] public string? Error { get; set; }
        [JsonPropertyName("version")] public string? Version { get; set; }
        [JsonPropertyName("files")] public BootstrapFiles Files { get; set; } = new();
    }

    public sealed class BootstrapFunction
    {
        private readonly ILogger<BootstrapFunction> _logger;
        private readonly IConfiguration _config;
        private readonly IHttpClientFactory _http;
        private static readonly JsonSerializerOptions _json = new(JsonSerializerDefaults.Web);
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, ConfigurationManager<OpenIdConnectConfiguration>> _cm = new();

        private const string PATH_NEW_HTML = "/z-wpsecure-cloud-sync__SYSTEM_DO_NOT_TOUCH/wpsecure_cloud_new.htm";
        private const string PATH_REPLY_HTML = "/z-wpsecure-cloud-sync__SYSTEM_DO_NOT_TOUCH/wpsecure_cloud_reply.htm";
        private const string PATH_NEW_TEXT = "/z-wpsecure-cloud-sync__SYSTEM_DO_NOT_TOUCH/wpsecure_cloud_new.txt";
        private const string PATH_REPLY_TEXT = "/z-wpsecure-cloud-sync__SYSTEM_DO_NOT_TOUCH/wpsecure_cloud_reply.txt";
        private const string PATH_APPT_NEW_HTML = "";
        private const string PATH_APPT_NEW_TEXT = "";

        public BootstrapFunction(ILogger<BootstrapFunction> logger, IConfiguration config, IHttpClientFactory http)
        {
            _logger = logger;
            _config = config;
            _http = http;
        }

        [Function("signatures-bootstrap")]
        public async Task<HttpResponseData> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "signatures/bootstrap")] FunctionsHttpRequestData req,
            FunctionContext ctx)
        {
            var res = req.CreateResponse(HttpStatusCode.OK);
            res.Headers.Add("Content-Type", "application/json; charset=utf-8");

            try
            {
                var body = await JsonSerializer.DeserializeAsync<BootstrapRequest>(req.Body, _json);
                if (body is null || string.IsNullOrWhiteSpace(body.IdToken))
                {
                    _logger.LogWarning("signatures-bootstrap: missing_id_token");
                    await res.WriteStringAsync(JsonSerializer.Serialize(new BootstrapResult { Success = false, Error = "missing_id_token", Version = null, Files = new BootstrapFiles() }, _json));
                    return res;
                }

                var tenantId = _config["Backend:TenantId"] ?? string.Empty;
                var backendClientId = _config["Backend:ClientId"] ?? string.Empty;
                var backendSecret = _config["Backend:ClientSecret"] ?? string.Empty;
                var authHost = _config["Cloud:AuthorityHost"] ?? "https://login.microsoftonline.com";
                var graphBase = _config["Graph:BaseUrl"] ?? "https://graph.microsoft.com";
                var scopes = _config["Graph:Scopes"];
                if (string.IsNullOrWhiteSpace(scopes)) scopes = graphBase.TrimEnd('/') + "/.default";

                var expectedAud1 = _config["Frontend:ClientId"];
                var expectedAud2 = _config["Frontend:AppIdUri"];

                if (string.IsNullOrWhiteSpace(tenantId) || string.IsNullOrWhiteSpace(backendClientId) || string.IsNullOrWhiteSpace(backendSecret))
                {
                    _logger.LogError("signatures-bootstrap: server_misconfigured");
                    await res.WriteStringAsync(JsonSerializer.Serialize(new BootstrapResult { Success = false, Error = "server_misconfigured", Version = null, Files = new BootstrapFiles() }, _json));
                    return res;
                }

                var valid = await ValidateUserTokenAsync(body.IdToken.Trim(), authHost, tenantId, expectedAud1, expectedAud2, _logger);
                if (!valid)
                {
                    await res.WriteStringAsync(JsonSerializer.Serialize(new BootstrapResult { Success = false, Error = "invalid_token", Version = null, Files = new BootstrapFiles() }, _json));
                    return res;
                }

                var authority = authHost.TrimEnd('/') + "/" + tenantId;
                var cca = ConfidentialClientApplicationBuilder
                    .Create(backendClientId)
                    .WithClientSecret(backendSecret)
                    .WithAuthority(authority)
                    .Build();

                var obo = await cca.AcquireTokenOnBehalfOf(new[] { scopes }, new UserAssertion(body.IdToken.Trim())).ExecuteAsync();

                var http = _http.CreateClient();
                http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", obo.AccessToken);

                var files = new BootstrapFiles
                {
                    NewHtml = await TryReadTextAsync(http, graphBase, PATH_NEW_HTML, _logger),
                    ReplyHtml = await TryReadTextAsync(http, graphBase, PATH_REPLY_HTML, _logger),
                    NewText = await TryReadTextAsync(http, graphBase, PATH_NEW_TEXT, _logger),
                    ReplyText = await TryReadTextAsync(http, graphBase, PATH_REPLY_TEXT, _logger),
                    ApptNewHtml = await TryReadTextAsync(http, graphBase, PATH_APPT_NEW_HTML, _logger),
                    ApptNewText = await TryReadTextAsync(http, graphBase, PATH_APPT_NEW_TEXT, _logger)
                };

                var result = new BootstrapResult { Success = true, Version = System.DateTime.UtcNow.ToString("yyyyMMdd-HHmmss"), Files = files };
                await res.WriteStringAsync(JsonSerializer.Serialize(result, _json));
                return res;
            }
            catch (JsonException je)
            {
                _logger.LogError(je, "signatures-bootstrap: invalid_json");
                await res.WriteStringAsync(JsonSerializer.Serialize(new BootstrapResult { Success = false, Error = "invalid_json", Version = null, Files = new BootstrapFiles() }, _json));
                return res;
            }
            catch (MsalServiceException msalEx)
            {
                _logger.LogError(msalEx, "signatures-bootstrap: obo_failed");
                await res.WriteStringAsync(JsonSerializer.Serialize(new BootstrapResult { Success = false, Error = "obo_failed", Version = null, Files = new BootstrapFiles() }, _json));
                return res;
            }
            catch (System.Exception ex)
            {
                _logger.LogError(ex, "signatures-bootstrap: internal_error");
                await res.WriteStringAsync(JsonSerializer.Serialize(new BootstrapResult { Success = false, Error = "internal_error", Version = null, Files = new BootstrapFiles() }, _json));
                return res;
            }
        }

        private static async Task<string?> TryReadTextAsync(System.Net.Http.HttpClient http, string graphBase, string path, ILogger logger)
        {
            if (string.IsNullOrWhiteSpace(path)) return null;
            try
            {
                var url = graphBase.TrimEnd('/') + "/v1.0/me/drive/root:" + path + ":/content";
                using var resp = await http.GetAsync(url);
                if (resp.StatusCode == HttpStatusCode.NotFound) return null;
                resp.EnsureSuccessStatusCode();
                return await resp.Content.ReadAsStringAsync();
            }
            catch (System.Net.Http.HttpRequestException hre)
            {
                logger.LogError(hre, "graph_read_failed");
                return null;
            }
        }

        private static async Task<bool> ValidateUserTokenAsync(string rawToken, string authorityHost, string tenantId, string? expectedAud1, string? expectedAud2, ILogger logger)
        {
            try
            {
                var issuer = authorityHost.TrimEnd('/') + "/" + tenantId + "/v2.0";
                var metadata = issuer + "/.well-known/openid-configuration";
                var cfgMgr = _cm.GetOrAdd(metadata, m => new ConfigurationManager<OpenIdConnectConfiguration>(m, new OpenIdConnectConfigurationRetriever(), new HttpDocumentRetriever { RequireHttps = true }));
                var cfg = await cfgMgr.GetConfigurationAsync(default);
                var audiences = new System.Collections.Generic.List<string>();
                if (!string.IsNullOrWhiteSpace(expectedAud1)) audiences.Add(expectedAud1!);
                if (!string.IsNullOrWhiteSpace(expectedAud2)) audiences.Add(expectedAud2!);
                var tvp = new TokenValidationParameters
                {
                    ValidateIssuer = true,
                    ValidIssuer = issuer,
                    ValidateIssuerSigningKey = true,
                    IssuerSigningKeys = cfg.SigningKeys,
                    ValidateLifetime = true,
                    ClockSkew = System.TimeSpan.FromMinutes(2),
                    ValidateAudience = audiences.Count > 0,
                    ValidAudiences = audiences.Count > 0 ? audiences : null
                };
                var handler = new System.IdentityModel.Tokens.Jwt.JwtSecurityTokenHandler();
                handler.ValidateToken(rawToken, tvp, out _);
                return true;
            }
            catch (System.Exception ex)
            {
                logger.LogError(ex, "token_validation_failed");
                return false;
            }
        }
    }
}

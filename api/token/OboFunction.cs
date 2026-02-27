
using System.Net;
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

namespace WPSecure.Api.Token
{
    public sealed class OboRequest
    {
        [JsonPropertyName("id_token")] public string? IdToken { get; set; }
    }

    public sealed class OboResult
    {
        [JsonPropertyName("success")] public bool Success { get; set; }
        [JsonPropertyName("error")] public string? Error { get; set; }
        [JsonPropertyName("access_token")] public string? AccessToken { get; set; }
    }

    public sealed class OboFunction
    {
        private readonly ILogger<OboFunction> _logger;
        private readonly IConfiguration _config;
        private static readonly JsonSerializerOptions _json = new(JsonSerializerDefaults.Web);
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, ConfigurationManager<OpenIdConnectConfiguration>> _cm = new();

        public OboFunction(ILogger<OboFunction> logger, IConfiguration config)
        {
            _logger = logger;
            _config = config;
        }

        [Function("token-obo")]
        public async Task<HttpResponseData> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "token/obo")] FunctionsHttpRequestData req,
            FunctionContext ctx)
        {
            var res = req.CreateResponse(HttpStatusCode.OK);
            res.Headers.Add("Content-Type", "application/json; charset=utf-8");

            try
            {
                var body = await JsonSerializer.DeserializeAsync<OboRequest>(req.Body, _json);
                if (body is null || string.IsNullOrWhiteSpace(body.IdToken))
                {
                    _logger.LogWarning("token-obo: missing_id_token");
                    await res.WriteStringAsync(JsonSerializer.Serialize(new OboResult { Success = false, Error = "missing_id_token" }, _json));
                    return res;
                }

                var tenantId = _config["BACKEND_TENANTID"] ?? string.Empty;
                var backendClientId = _config["BACKEND_CLIENTID"] ?? string.Empty;
                var backendSecret = _config["BACKEND_CLIENTSECRET"] ?? string.Empty;
                var authHost = _config["CLOUD_AUTHORITYHOST"] ?? "https://login.microsoftonline.com";
                var expectedAud1 = _config["FRONTEND_CLIENTID"];
                var expectedAud2 = _config["FRONTEND_APPIDURI"];
                var graphBase = _config["GRAPH_BASEURL"] ?? "https://graph.microsoft.com";
                var scopes = _config["GRAPH_SCOPES"];
                if (string.IsNullOrWhiteSpace(scopes)) scopes = graphBase.TrimEnd('/') + "/.default";

                if (string.IsNullOrWhiteSpace(tenantId) || string.IsNullOrWhiteSpace(backendClientId) || string.IsNullOrWhiteSpace(backendSecret))
                {
                    _logger.LogError("token-obo: server_misconfigured");
                    await res.WriteStringAsync(JsonSerializer.Serialize(new OboResult { Success = false, Error = "server_misconfigured" }, _json));
                    return res;
                }

                var valid = await ValidateUserTokenAsync(body.IdToken.Trim(), authHost, tenantId, expectedAud1, expectedAud2, _logger);
                if (!valid)
                {
                    await res.WriteStringAsync(JsonSerializer.Serialize(new OboResult { Success = false, Error = "invalid_token" }, _json));
                    return res;
                }

                var authority = authHost.TrimEnd('/') + "/" + tenantId;
                var cca = ConfidentialClientApplicationBuilder
                    .Create(backendClientId)
                    .WithClientSecret(backendSecret)
                    .WithAuthority(authority)
                    .Build();

                try
                {
                    var obo = await cca.AcquireTokenOnBehalfOf(new[] { scopes }, new UserAssertion(body.IdToken.Trim())).ExecuteAsync();
                    await res.WriteStringAsync(JsonSerializer.Serialize(new OboResult { Success = true, AccessToken = obo.AccessToken }, _json));
                    return res;
                }
                catch (MsalServiceException msalEx)
                {
                    _logger.LogError(msalEx, "token-obo: obo_failed");
                    await res.WriteStringAsync(JsonSerializer.Serialize(new OboResult { Success = false, Error = "obo_failed" }, _json));
                    return res;
                }
            }
            catch (JsonException je)
            {
                _logger.LogError(je, "token-obo: invalid_json");
                await res.WriteStringAsync(JsonSerializer.Serialize(new OboResult { Success = false, Error = "invalid_json" }, _json));
                return res;
            }
            catch (System.Exception ex)
            {
                _logger.LogError(ex, "token-obo: internal_error");
                await res.WriteStringAsync(JsonSerializer.Serialize(new OboResult { Success = false, Error = "internal_error" }, _json));
                return res;
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
                logger.LogError(ex, "token-obo: token_validation_failed");
                return false;
            }
        }
    }
}

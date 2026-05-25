using System.Net;
using System.Text.Json;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Azure.Identity;

namespace UpceGym.Functions
{
    public class RecordPassage
    {
        private readonly ILogger<RecordPassage> _logger;
        private static GraphServiceClient? _graphClient;

        // Načtení konfigurace z ENV
        private static readonly string TenantId = Environment.GetEnvironmentVariable("AzureAd_TenantId") ?? string.Empty;
        private static readonly string ClientId = Environment.GetEnvironmentVariable("AzureAd_ClientId") ?? string.Empty;
        private static readonly string ClientSecret = Environment.GetEnvironmentVariable("AzureAd_ClientSecret") ?? string.Empty;
        private static readonly string SiteId = Environment.GetEnvironmentVariable("SharePoint_SiteId") ?? string.Empty;
        private static readonly string VydanaPermanentkaListId = Environment.GetEnvironmentVariable("SharePoint_VydanaPermanentkaListId") ?? string.Empty;
        private static readonly string NavstevaListId = Environment.GetEnvironmentVariable("SharePoint_NavstevaListId") ?? string.Empty;
        private static readonly string ClenstviListId = Environment.GetEnvironmentVariable("SharePoint_ClenstviListId") ?? string.Empty;

        public RecordPassage(ILogger<RecordPassage> logger)
        {
            _logger = logger;
            InitializeGraphClient();
        }

        private void InitializeGraphClient()
        {
            if (_graphClient == null)
            {
                var credential = new ClientSecretCredential(TenantId, ClientId, ClientSecret);
                _graphClient = new GraphServiceClient(credential);
            }
        }

        [Function("RecordPassage")]
        public async Task<HttpResponseData> Run([HttpTrigger(AuthorizationLevel.Function, "post", Route = "gym/passage")] HttpRequestData req)
        {
            _logger.LogInformation("Zpracování fyzického průchodu turniketem.");

            // Načtení a parsování těla požadavku (JSON)
            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            using var doc = JsonDocument.Parse(requestBody);
            var root = doc.RootElement;

            // Kontrola povinného parametru 'Smer'
            if (!root.TryGetProperty("Smer", out var smerProp))
            {
                var badResponse = req.CreateResponse(HttpStatusCode.BadRequest);
                await badResponse.WriteAsJsonAsync(new { success = false, message = "Chybí povinný parametr Smer." });
                return badResponse;
            }

            string smer = smerProp.ToString(); // "Prichod" nebo "Odchod"

            try
            {
                // ==========================================
                // SCÉNÁŘ A: PŘÍCHOD (Stateful z VerifyCard)
                // ==========================================
                if (smer.Equals("Prichod", StringComparison.OrdinalIgnoreCase))
                {
                    if (!root.TryGetProperty("ClenID", out var clenIdProp))
                    {
                        var badResponse = req.CreateResponse(HttpStatusCode.BadRequest);
                        await badResponse.WriteAsJsonAsync(new { success = false, message = "Pro příchod je vyžadován parametr ClenID." });
                        return badResponse;
                    }

                    string clenId = clenIdProp.ToString();
                    root.TryGetProperty("PermanentkaID", out var permIdProp);
                    string permId = permIdProp.ToString();

                    _logger.LogInformation($"Zaznamenávám PŘÍCHOD. ČlenID: {clenId}, PermanentkaID: {permId}");

                    // KROK A.1: Odečtení vstupu z permanentky (HTTP PATCH)
                    if (!string.IsNullOrEmpty(permId))
                    {
                        // Načtení aktuálního počtu vstupů ze SharePointu
                        var permice = await _graphClient!.Sites[SiteId].Lists[VydanaPermanentkaListId].Items[permId]
                            .GetAsync(rc => rc.QueryParameters.Expand = new[] { "fields" });

                        if (permice?.Fields?.AdditionalData.TryGetValue("PocetZbyvajicichVstupu", out var aktualniVstupyObj) == true &&
                            double.TryParse(aktualniVstupyObj?.ToString(), out double aktualniVstupy))
                        {
                            int novyPocet = (int)aktualniVstupy - 1;
                            if (novyPocet < 0) novyPocet = 0;

                            // Příprava dat pro aktualizaci řádku
                            var updateFields = new ListItem
                            {
                                Fields = new FieldValueSet
                                {
                                    AdditionalData = new Dictionary<string, object>
                                    {
                                        { "PocetZbyvajicichVstupu", novyPocet }
                                    }
                                }
                            };

                            // Odeslání změn zpět do SharePointu (PATCH)
                            await _graphClient.Sites[SiteId].Lists[VydanaPermanentkaListId].Items[permId].PatchAsync(updateFields);
                            _logger.LogInformation($"Vstup úspěšně odečten. Nový počet zbývajících vstupů: {novyPocet}");
                        }
                    }

                    // KROK A.2: Zápis nového řádku návštěvy (HTTP POST)
                    var novaNavsteva = new ListItem
                    {
                        Fields = new FieldValueSet
                        {
                            AdditionalData = new Dictionary<string, object>
                            {
                                { "Title", $"Příchod člena ID {clenId}" },
                                { "ClenID", int.Parse(clenId) },
                                { "VydanaPermanentkaID", !string.IsNullOrEmpty(permId) ? int.Parse(permId) : 0 },
                                { "CasPrichodu", DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ") } // ISO formát pro SharePoint
                            }
                        }
                    };

                    await _graphClient!.Sites[SiteId].Lists[NavstevaListId].Items.PostAsync(novaNavsteva);
                    _logger.LogInformation("Nový záznam návštěvy (příchod) byl úspěšně vytvořen.");
                }
                
                // ==========================================
                // SCÉNÁŘ B: ODCHOD (Bezstavový podle NfcID)
                // ==========================================
                else if (smer.Equals("Odchod", StringComparison.OrdinalIgnoreCase))
                {
                    if (!root.TryGetProperty("NfcID", out var nfcIdProp))
                    {
                        var badResponse = req.CreateResponse(HttpStatusCode.BadRequest);
                        await badResponse.WriteAsJsonAsync(new { success = false, message = "Pro odchod je vyžadován parametr NfcID." });
                        return badResponse;
                    }

                    string nfcId = nfcIdProp.ToString();
                    _logger.LogInformation($"Zaznamenávám ODCHOD pro NFC kartu: {nfcId}");

                    // KROK B.1: Dohledání ClenID v seznamu Členství pomocí NfcID
                    string clenFilter = $"fields/NfcID eq '{nfcId}'";
                    var clenoveResult = await _graphClient!.Sites[SiteId].Lists[ClenstviListId].Items
                        .GetAsync(rc => 
                        {
                            rc.QueryParameters.Filter = clenFilter;
                            rc.QueryParameters.Expand = new[] { "fields" };
                        });

                    var clenItem = clenoveResult?.Value?.FirstOrDefault();
                    if (clenItem == null)
                    {
                        _logger.LogWarning($"Při odchodu nebyl nalezen žádný člen s kartou {nfcId}.");
                        var notFoundResponse = req.CreateResponse(HttpStatusCode.NotFound);
                        await notFoundResponse.WriteAsJsonAsync(new { success = false, message = "Člen s touto kartou neexistuje." });
                        return notFoundResponse;
                    }

                    string clenIdZkarty = clenItem.Id ?? string.Empty;
                    _logger.LogInformation($"Karta {nfcId} úspěšně spárována s ČlenID: {clenIdZkarty}. Hledám otevřenou návštěvu...");

                    // KROK B.2: Vyhledání otevřené návštěvy (kde chybí CasOdchodu)
                    string filter = $"fields/ClenID eq {clenIdZkarty} and fields/CasOdchodu eq null";
                    var navstevyResult = await _graphClient!.Sites[SiteId].Lists[NavstevaListId].Items
                        .GetAsync(rc => 
                        {
                            rc.QueryParameters.Filter = filter;
                            rc.QueryParameters.Expand = new[] { "fields" };
                        });

                    var posledniOtevrenaNavsteva = navstevyResult?.Value?.FirstOrDefault();

                    // KROK B.3: Aktualizace času odchodu (HTTP PATCH) nebo Nouzový odchod
                    if (posledniOtevrenaNavsteva != null)
                    {
                        var updateNavsteva = new ListItem
                        {
                            Fields = new FieldValueSet
                            {
                                AdditionalData = new Dictionary<string, object>
                                {
                                    { "CasOdchodu", DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ") }
                                }
                            }
                        };

                        await _graphClient.Sites[SiteId].Lists[NavstevaListId].Items[posledniOtevrenaNavsteva.Id].PatchAsync(updateNavsteva);
                        _logger.LogInformation($"Čas odchodu úspěšně zapsán do návštěvy ID: {posledniOtevrenaNavsteva.Id}");
                    }
                    else
                    {
                        // Fail-safe mechanismus: Uživatel je uvnitř, ale záznam o příchodu chybí -> vytvoříme řádek rovnou s odchodem
                        _logger.LogWarning($"Nenalezena žádná otevřená návštěva pro člena {clenIdZkarty}. Zapisuji nouzový odchod.");
                        
                        var nouzovyOdchod = new ListItem
                        {
                            Fields = new FieldValueSet
                            {
                                AdditionalData = new Dictionary<string, object>
                                {
                                    { "Title", $"Nouzový odchod člena ID {clenIdZkarty}" },
                                    { "ClenID", int.Parse(clenIdZkarty) },
                                    { "CasOdchodu", DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ") }
                                }
                            }
                        };
                        await _graphClient.Sites[SiteId].Lists[NavstevaListId].Items.PostAsync(nouzovyOdchod);
                    }
                }

                // Finální úspěšná odpověď pro hardware turniketu
                var response = req.CreateResponse(HttpStatusCode.OK);
                await response.WriteAsJsonAsync(new { success = true, message = "Průchod úspěšně zaznamenán v databázi." });
                return response;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Kritická chyba při zápisu průchodu: {ex.Message}");
                var errorResponse = req.CreateResponse(HttpStatusCode.InternalServerError);
                await errorResponse.WriteAsJsonAsync(new { success = false, message = "Vnitřní chyba serveru při zápisu průchodu." });
                return errorResponse;
            }
        }
    }
}
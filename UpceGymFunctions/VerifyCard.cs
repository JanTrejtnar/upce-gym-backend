using System.Net;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using System.Text.Json;
using Microsoft.Graph;

namespace UpceGym.Functions
{
    public class VerifyCard
    {
        private readonly ILogger<VerifyCard> _logger;
        private readonly GraphServiceClient _graphClient;
        
        private static readonly string SiteId = Environment.GetEnvironmentVariable("SharePoint_SiteId") ?? string.Empty;
        private static readonly string ClenstviListId = Environment.GetEnvironmentVariable("SharePoint_ClenstviListId") ?? string.Empty;
        private static readonly string VydanaPermanentkaListId = Environment.GetEnvironmentVariable("SharePoint_VydanaPermanentkaListId") ?? string.Empty;

        public VerifyCard(ILogger<VerifyCard> logger, GraphServiceClient graphClient)
        {
            _logger = logger;
            _graphClient = graphClient; // GraphServiceClient injektován přes DI a je připraven k použití
        }

        [Function("VerifyCard")]
        public async Task<HttpResponseData> Run(
            [HttpTrigger(AuthorizationLevel.Function, "post", Route = "gym/verify")] HttpRequestData req)
        {
            _logger.LogInformation("Zpracování požadavku na vstup do posilovny.");

            // 1. Načtení NFC ID z požadavku turniketu
            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            var cardData = JsonSerializer.Deserialize<CardRequest>(requestBody);

            var response = req.CreateResponse();

            if (cardData == null || string.IsNullOrEmpty(cardData.NfcId))
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                await response.WriteAsJsonAsync(new { allowed = false, message = "Chybí NfcID v požadavku." });
                return response;
            }

            try
            {
                // 2. KROK: Hledání člena v seznamu 'Členství' podle sloupce 'NfcID' - OData filtr
                var clenListItems = await _graphClient.Sites[SiteId].Lists[ClenstviListId].Items
                    .GetAsync(requestConfiguration => {
                        requestConfiguration.QueryParameters.Filter = $"fields/NfcID eq '{cardData.NfcId}'";
                        requestConfiguration.QueryParameters.Expand = new string[] { "fields" };
                    });

                var clenItem = clenListItems?.Value?.FirstOrDefault();

                if (clenItem == null)
                {
                    _logger.LogWarning($"Neznámé NFC ID: {cardData.NfcId}");
                    response.StatusCode = HttpStatusCode.OK;
                    await response.WriteAsJsonAsync(new { allowed = false, message = "Neznámá karta / Člen nenalezen." });
                    return response;
                }

                // Získání ID člena (ID řádku ze seznamu Členství)
                string clenId = clenItem.Id ?? String.Empty;
                string jmenoClena = clenItem.Fields?.AdditionalData.ContainsKey("Title") == true
                    ? clenItem.Fields.AdditionalData["Title"]?.ToString() ?? "Neznámý člen"
                    : "Neznámý člen";

                _logger.LogInformation($"Člen nalezen: {jmenoClena} (ClenID: {clenId}). Ověřuji permanentku...");

                // 3. KROK: Hledání permanentky v seznamu 'VydanaPermanentka' podle 'ClenID'
                var permListItems = await _graphClient.Sites[SiteId].Lists[VydanaPermanentkaListId].Items
                    .GetAsync(requestConfiguration => {
                        requestConfiguration.QueryParameters.Filter = $"fields/ClenID eq '{clenId}' and fields/Stav eq 'Aktivní'";
                        requestConfiguration.QueryParameters.Expand = new string[] { "fields" };
                    });

                var permanentky = permListItems?.Value;

                if (permanentky == null || !permanentky.Any())
                {
                    _logger.LogWarning($"Člen {jmenoClena} nemá žádnou vydanou permanentku.");
                    response.StatusCode = HttpStatusCode.OK;
                    await response.WriteAsJsonAsync(new { allowed = false, message = "Nemáte žádnou permanentku." });
                    return response;
                }

                // 4. KROK: Validace byznys logiky nad permanentkami
                bool maPlatnyVstup = false;
                string duvodZamitnuti = "Žádná z permanentek nesplňuje podmínky vstupu.";
                string platnaPermiceId = string.Empty;
                DateTime dnes = DateTime.Today;

                _logger.LogInformation($"Počet nalezených aktivních permanentek pro člena v databázi: {permanentky.Count}");

                foreach (var perm in permanentky)
                {
                    var fields = perm.Fields?.AdditionalData;

                    // Načtení hodnot ze SharePoint sloupců
                    string stav = fields!.ContainsKey("Stav") ? fields["Stav"]?.ToString() ?? string.Empty : string.Empty;
                    bool jeZaplacena = fields.ContainsKey("JeZaplacena") && Convert.ToBoolean(fields["JeZaplacena"]);
                    
                    int pocetVstupu = 0;
                    if (fields.ContainsKey("PocetZbyvajicichVstupu"))
                    {
                        string surovaHodnota = fields["PocetZbyvajicichVstupu"]?.ToString() ?? "0";
                        
                        if (double.TryParse(surovaHodnota, out double desetinneCislo))
                        {
                            // Převedeme (zaokrouhlíme dolů) double na int
                            pocetVstupu = (int)desetinneCislo;
                        }
                        else
                        {
                            _logger.LogWarning($"Nelze převést hodnotu 'PocetZbyvajicichVstupu' na číslo: {surovaHodnota}. Nastavuji počet vstupů na 0.");
                            pocetVstupu = 0; // Pokud převod selže, nastavíme počet vstupů na 0
                        }
                    }

                    DateTime platnostOd = fields.ContainsKey("PlatnostOd") ? Convert.ToDateTime(fields["PlatnostOd"]) : DateTime.MaxValue;
                    DateTime platnostDo = fields.ContainsKey("PlatnostDo") ? Convert.ToDateTime(fields["PlatnostDo"]) : DateTime.MinValue;

                    _logger.LogInformation($"Nalezena permanentka. Stav permanentky: {stav}. Je zaplacena?: {jeZaplacena}.Platnost: {platnostOd} - {platnostDo}. Počet zbývajících vstupů: {pocetVstupu}. Ověřuji permanentku...");

                    // Kontrola podmínek:
                    // - Stav == "Aktivní"
                    // - JeZaplacena == true
                    // - Dnešní datum je v rozsahu PlatnostOd až PlatnostDo
                    // - Počet zbývajících vstupů > 0
                    if (stav == "Aktivní" && jeZaplacena && dnes >= platnostOd && dnes <= platnostDo && pocetVstupu > 0)
                    {
                        maPlatnyVstup = true;
                        platnaPermiceId = perm.Id ?? string.Empty; // Uložíme ID řádku permanentky, ze které bude odečítán vstup
                        break; // Našli jsme platnou permanentku, končíme cyklus
                    }
                }

                // 5. KROK: Finální odpověď turniketu
                if (maPlatnyVstup)
                {
                    _logger.LogInformation($"Vstup POVOLEN pro člena {jmenoClena}.");
                    
                    var successResponse = req.CreateResponse(HttpStatusCode.OK);
                    
                    // Do JSON odpovědi přibalíme data, která si bezstavový turniket na chvíli uloží do paměti
                    await successResponse.WriteAsJsonAsync(new { 
                        allowed = true, 
                        message = $"Vítej, {jmenoClena}! Vstup povolen.",
                        clenId = clenId,
                        permanentkaId = platnaPermiceId
                    });
                    
                    return successResponse;
                }
                else
                {
                    _logger.LogWarning($"Vstup ZAMÍTNUT pro člena {jmenoClena}. Podmínky permanentky nesplněny.");
                    response.StatusCode = HttpStatusCode.OK;
                    await response.WriteAsJsonAsync(new { allowed = false, message = duvodZamitnuti });
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Chyba při komunikaci se SharePointem: {ex.Message}");
                
                var errorResponse = req.CreateResponse(HttpStatusCode.InternalServerError);
                await errorResponse.WriteAsJsonAsync(new { allowed = false, message = "Chyba na straně serveru při ověřování v databázi." });
                return errorResponse;
            }

            return response;
        }
    }

    public class CardRequest
    {
        public string NfcId { get; set; } = string.Empty;
    }
}
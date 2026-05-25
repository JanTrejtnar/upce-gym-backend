# UpceGymManagement – Řídicí API pro fitness turnikety

Tento projekt obsahuje backendové řešení pro autonomní správu fitness turniketů postavené na architektuře **Azure Functions (.NET 8 Isolated Worker)**. Systém komunikuje asynchronně s **Microsoft Graph API**, jako datové úložiště využívá **SharePoint Online** (seznamy pro Členství, Permanentky a Návštěvy) a je navržen podle principů událostmi řízené architektury (*Event-Driven Architecture*).

## Architektura systému

1. **VerifyCard (Ověření identity):** Turniket pošle surové NFC ID. Funkce ověří existenci člena a platnost jeho permanentky v SharePointu. Pokud je vše v pořádku, vrátí `allowed: true` spolu s ID člena a ID permanentky. Turniket se fyzicky odemkne a tato ID si krátkodobě uloží do paměti.
2. **RecordPassage (Zápis průchodu):** Jakmile senzory v turniketu zaznamenají reálné otočení ramene, turniket bezstavově odešle požadavek na zápis. Při příchodu se odečte 1 vstup a založí se záznam o návštěvě. Při odchodu (aktivovaném na vnitřní čtečce) systém podle NFC ID dohledá člena, najde otevřenou návštěvu a zapíše čas odchodu.

---

## Lokální testování (Příkazový řádek CMD ve Windows)

Pro lokální testování je nutné mít spuštěné Azure Functions Core Tools pomocí příkazu:
```bash
func start
```

Vzhledem k specifickému parsování uvozovek v klasickém příkazovém řádku Windows (CMD) je nutné testovací curl příkazy posílat přesně v níže uvedených formátech (s escapováním \").

### Krok 1: Ověření karty na venkovní čtečce (Příchod)
Simulace situace, kdy uživatel přiloží kartu k turniketu před vstupem do fitness centra.

```bash
curl -X POST http://localhost:7071/api/gym/verify -H "Content-Type: application/json" \ -d "{\"NfcId\":\"TEST123456\"}"
```

Očekávaná odpověď:
```json
{
  "allowed": true,
  "message": "Vítej, Jan Trejtnar! Vstup povolen.",
  "clenId": "1",
  "permanentkaId": "12"
}
```

### Krok 2: Zaznamenání fyzického PŘÍCHODU (Odečtení vstupu)
Simulace momentu, kdy se rameno turniketu fyzicky otočilo směrem dovnitř.
(V příkazu níže nahraďte 12 reálným ID permanentky z Kroku 1).

```bash
curl -X POST http://localhost:7071/api/gym/passage -H "Content-Type: application/json" \ -d "{\"ClenID\":\"1\",\"PermanentkaID\":\"12\",\"Smer\":\"Prichod\"}"
```

Výsledek v SharePointu:
1. V seznamu Permanentky klesne pole PocetZbyvajicichVstupu o 1.
2. V seznamu Navstevy se vytvoří nový řádek s vyplněným polem CasPrichodu (pole CasOdchodu zůstává prázdné).

### Krok 3: Zaznamenání fyzického ODCHODU (Uzavření návštěvy)
Simulace odchodu cvičence domů přes vnitřní čtečku turniketu. Hardware funguje bezstavově, posílá pouze surové NFC ID a směr.

```bash
curl -X POST http://localhost:7071/api/gym/passage -H "Content-Type: application/json" \ -d "{\"NfcID\":\"TEST123456\",\"Smer\":\"Odchod\"}"
```

Výsledek v SharePointu:
1. Systém automaticky dohledá člena, najde jeho otevřený záznam v seznamu Navstevy a bezpečně do něj dopíše aktuální UTC čas do pole CasOdchodu.

## Konfigurace prostředí
Pro správný chod aplikace je nutné mít v kořenovém adresáři nastavené proměnné prostředí pro připojení k Azure AD (Entra ID) a SharePointu:
```json
{
  "IsEncrypted": false,
  "Values": {
    "AzureWebJobsStorage": "UseDevelopmentStorage=true",
    "FUNCTIONS_WORKER_RUNTIME": "dotnet-isolated",
    "AzureAd_TenantId": "VAŠE_TENANT_ID",
    "AzureAd_ClientId": "VAŠE_CLIENT_ID",
    "AzureAd_ClientSecret": "VÁŠ_CLIENT_SECRET",
    "SharePoint_SiteId": "VAŠE_SHAREPOINT_SITE_ID",
    "SharePoint_ClenstviListId": "GUID_SEZNAMU_CLENSTVI",
    "SharePoint_PermanentkyListId": "GUID_SEZNAMU_PERMANENTEK",
    "SharePoint_NavstevyListId": "GUID_SEZNAMU_NAVSTEV"
  }
}
```

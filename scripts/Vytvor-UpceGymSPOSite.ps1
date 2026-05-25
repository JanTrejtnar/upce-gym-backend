# 1. Připojení k prostředí SharePoint Online
Connect-PnPOnline -Url "https://legranc.sharepoint.com" -UseWebLogin

# 2. Vytvoření nového komunikačního webu
$siteUrl = New-PnPSite -Type CommunicationSite -Title "KTS - Gym Management" -Url "KTSGymManagement"

# 3. Připojení k nově vytvořenému webu
Connect-PnPOnline -Url "https://legranc.sharepoint.com/sites/UPCEGymManagement" -UseWebLogin #$siteUrl -UseWebLogin


# 4. Vytvoření seznamů

# =========================================================================
# 4.1 Seznam - "Člen"
# =========================================================================
$clenListTitle = "Člen"
$clenListName = "Clen"

Write-Host "Vytvářím seznam $clenListTitle..." -ForegroundColor Cyan
New-PnPList -Title $clenListTitle -Template GenericList -Url $clenListName

# 4.1.a. Vytvoření sloupců
Add-PnPField -List $clenListTitle -DisplayName "NetID" -InternalName "NetID" -Type Text -AddToDefaultView
Add-PnPField -List $clenListTitle -DisplayName "Jméno" -InternalName "Jmeno" -Type Text -AddToDefaultView
Add-PnPField -List $clenListTitle -DisplayName "Příjmení" -InternalName "Prijmeni" -Type Text -AddToDefaultView
Add-PnPField -List $clenListTitle -DisplayName "Email" -InternalName "Email" -Type Text -AddToDefaultView
Add-PnPField -List $clenListTitle -DisplayName "Telefon" -InternalName "Telefon" -Type Text -AddToDefaultView
Add-PnPField -List $clenListTitle -DisplayName "NfcID" -InternalName "NfcID" -Type Text -AddToDefaultView
Add-PnPField -List $clenListTitle -DisplayName "Datum registrace" -InternalName "DatumRegistrace" -Type DateTime -AddToDefaultView
Add-PnPField -List $clenListTitle -DisplayName "Poznámka" -InternalName "Poznamka" -Type Note -AddToDefaultView

# 4.1.b. Nastavení unikátnosti pro sloupec NetID
Write-Host "Nastavuji sloupec NetID a NfcID jako unikátní..." -ForegroundColor Yellow

# Načtení pole
$netIdField = Get-PnPField -List $clenListTitle -Identity "NetID"
$nfcIdField = Get-PnPField -List $clenListTitle -Identity "NfcID"

# SharePoint vyžaduje, aby unikátní sloupec byl zároveň indexovaný
$netIdField.Indexed = $true
$nfcIdField.Indexed = $true

$netIdField.EnforceUniqueValues = $true
$nfcIdField.EnforceUniqueValues = $true

# Uložení změn do SharePointu
$netIdField.Update()
$nfcIdField.Update()
Invoke-PnPQuery

Write-Host "Seznam $clenListTitle a jeho pole byly úspěšně vytvořeny!" -ForegroundColor Green



# =========================================================================
# 4.2 Seznam - "Typ permanentky"
# =========================================================================
$typPermanentkyListTitle = "Typ permanentky"
$typPermanentkyListName = "TypPermanentky"

Write-Host "Vytvářím seznam $typPermanentkyListTitle..." -ForegroundColor Cyan
New-PnPList -Title $typPermanentkyListTitle -Template GenericList -Url $typPermanentkyListName

# 4.2.a. Vytvoření sloupců
Add-PnPField -List $typPermanentkyListTitle -DisplayName "Název" -InternalName "Nazev" -Type Text -AddToDefaultView
Add-PnPField -List $typPermanentkyListTitle -DisplayName "Cena" -InternalName "Cena" -Type Currency -AddToDefaultView
Add-PnPField -List $typPermanentkyListTitle -DisplayName "Počet možných vstupů" -InternalName "PocetMoznychVstupu" -Type Number -AddToDefaultView

Write-Host "Seznam $typPermanentkyListTitle a jeho pole byly úspěšně vytvořeny!" -ForegroundColor Green



# =========================================================================
# 4.3 Seznam - "Vydaná permanentka"
# =========================================================================
$vydanaPermanentkaListTitle = "Vydaná permanentka"
$vydanaPermanentkaListName = "VydanaPermanentka"

Write-Host "Vytvářím seznam $vydanaPermanentkaListTitle..." -ForegroundColor Cyan
New-PnPList -Title $vydanaPermanentkaListTitle -Template GenericList -Url $vydanaPermanentkaListName

# 4.3.a. Vytvoření sloupců
Add-PnPField -List $vydanaPermanentkaListTitle -DisplayName "Platnost od" -InternalName "PlatnostOd" -Type DateTime -AddToDefaultView
Add-PnPField -List $vydanaPermanentkaListTitle -DisplayName "Platnost do" -InternalName "PlatnostDo" -Type DateTime -AddToDefaultView
Add-PnPField -List $vydanaPermanentkaListTitle -DisplayName "Je zaplacena" -InternalName "JeZaplacena" -Type Boolean -AddToDefaultView
Add-PnPField -List $vydanaPermanentkaListTitle -DisplayName "Stav" -InternalName "Stav" -Type Choice -Choices "Aktivní","Expirovaná","Blokovaná, Neaktivní" -AddToDefaultView
Add-PnPField -List $vydanaPermanentkaListTitle -DisplayName "Platební metoda" -InternalName "PlatebniMetoda" -Type Choice -Choices "Hotovost","Bankovní převod","Karta" -AddToDefaultView
#Add-PnPField -List $vydanaPermanentkaListTitle -DisplayName "Člen" -InternalName "ClenLookup" -Type Lookup -AddToDefaultView
Add-PnPField -List $vydanaPermanentkaListTitle -DisplayName "NetID" -InternalName "NetID" -Type Text -AddToDefaultView


Add-PnPField -List $vydanaPermanentkaListTitle -DisplayName "Člen ID" -InternalName "ClenID" -Type Number -AddToDefaultView
#Add-PnPField -List $vydanaPermanentkaListTitle -DisplayName "Typ permanentky" -InternalName "TypPermanentkyLookup" -Type Lookup -AddToDefaultView
Add-PnPField -List $vydanaPermanentkaListTitle -DisplayName "Typ permanentky ID" -InternalName "TypPermanentkyID" -Type Number -AddToDefaultView

# Duplicitní sloupce - Typ Permanentky
Add-PnPField -List $vydanaPermanentkaListTitle -DisplayName "Počet možných vstupů" -InternalName "PocetMoznychVstupu" -Type Number -AddToDefaultView
Add-PnPField -List $vydanaPermanentkaListTitle -DisplayName "Počet zbývajících vstupů" -InternalName "PocetZbyvajicichVstupu" -Type Number -AddToDefaultView
Add-PnPField -List $vydanaPermanentkaListTitle -DisplayName "Typ permanentky" -InternalName "TypPermanentky" -Type Text -AddToDefaultView
Add-PnPField -List $vydanaPermanentkaListTitle -DisplayName "Cena" -InternalName "Cena" -Type Currency -AddToDefaultView

# 4.3.b. Nastavení indexovaných sloupců
Write-Host "Nastavuji indexované sloupce..." -ForegroundColor Yellow

# Načtení pole
$clenIdField = Get-PnPField -List $vydanaPermanentkaListTitle -Identity "ClenID"
$stavField = Get-PnPField -List $vydanaPermanentkaListTitle -Identity "Stav"

# Nastavení pole jako indexované
$clenIdField.Indexed = $true
$stavField.Indexed = $true

# Uložení změn do SharePointu
$clenIdField.Update()
$stavField.Update()
Invoke-PnPQuery

Write-Host "Seznam $vydanaPermanentkaListTitle a jeho pole byly úspěšně vytvořeny!" -ForegroundColor Green


# =========================================================================
# 4.4 Seznam - "Návštěva"
# =========================================================================
$navstevaListTitle = "Návštěva"
$navstevaListName = "Navsteva"

Write-Host "Vytvářím seznam $navstevaListTitle..." -ForegroundColor Cyan
New-PnPList -Title $navstevaListTitle -Template GenericList -Url $navstevaListName

# 4.4.a Vytvoření sloupců
Add-PnPField -List $navstevaListTitle -DisplayName "Čas příchodu" -InternalName "CasPrichodu" -Type DateTime -AddToDefaultView
Add-PnPField -List $navstevaListTitle -DisplayName "Čas odchodu" -InternalName "CasOdchodu" -Type DateTime -AddToDefaultView
Add-PnPField -List $navstevaListTitle -DisplayName "Vydaná permanentka ID" -InternalName "VydanaPermanentkaID" -Type Number -AddToDefaultView
Add-PnPField -List $navstevaListTitle -DisplayName "Člen ID" -InternalName "ClenID" -Type Number -AddToDefaultView

# 4.4.b. Nastavení indexovaných sloupců
Write-Host "Nastavuji indexované sloupce..." -ForegroundColor Yellow

# Načtení pole
$clenIdField = Get-PnPField -List $navstevaListTitle -Identity "ClenID"
$vydanaPermanentkaIdField = Get-PnPField -List $navstevaListTitle -Identity "VydanaPermanentkaID"
$casPrichoduField = Get-PnPField -List $navstevaListTitle -Identity "CasPrichodu"
$casOdchoduField = Get-PnPField -List $navstevaListTitle -Identity "CasOdchodu"


# Nastavení pole jako indexované
$clenIdField.Indexed = $true
$vydanaPermanentkaIdField.Indexed = $true
$casPrichoduField.Indexed = $true
$casOdchoduField.Indexed = $true

# Uložení změn do SharePointu
$clenIdField.Update()
$vydanaPermanentkaIdField.Update()
$casPrichoduField.Update()
$casOdchoduField.Update()
Invoke-PnPQuery

Write-Host "Seznam $navstevaListTitle a jeho pole byly úspěšně vytvořeny!" -ForegroundColor Green


# =========================================================================
# 4.5 Seznam - "Error Log"
# =========================================================================
$errorLogListTitle = "Error Log"
$errorLogListName = "Error Log"

Write-Host "Vytvářím seznam $errorLogListTitle..." -ForegroundColor Cyan
New-PnPList -Title $errorLogListTitle -Template GenericList -Url $errorLogListName

# 4.5.a Vytvoření sloupců
Add-PnPField -List $errorLogListTitle -DisplayName "Notifikace" -InternalName "Notifikace" -Type Note -AddToDefaultView
Add-PnPField -List $errorLogListTitle -DisplayName "Detail" -InternalName "Detail" -Type Note -AddToDefaultView
Add-PnPField -List $errorLogListTitle -DisplayName "Zdroj" -InternalName "Zdroj" -Type Choice -Choices "Power Apps","Power Automate" -AddToDefaultView
Add-PnPField -List $errorLogListTitle -DisplayName "Url" -InternalName "Url" -Type URL -AddToDefaultView

Add-PnPField -List $errorLogListTitle -DisplayName "Člen ID" -InternalName "ClenID" -Type Number -AddToDefaultView
Add-PnPField -List $errorLogListTitle -DisplayName "Vydaná permanentka ID" -InternalName "VydanaPermanentkaID" -Type Number -AddToDefaultView
Add-PnPField -List $errorLogListTitle -DisplayName "Návštěva ID" -InternalName "NavstevaID" -Type Number -AddToDefaultView
Add-PnPField -List $errorLogListTitle -DisplayName "Závada ID" -InternalName "ZavadaID" -Type Number -AddToDefaultView

Write-Host "Seznam $errorLogListTitle a jeho pole byly úspěšně vytvořeny!" -ForegroundColor Green



# =========================================================================
# 4.6 Seznam - "Audit Log"
# =========================================================================
$auditLogListTitle = "Audit Log"
$auditLogListName = "AuditLog"

Write-Host "Vytvářím seznam $auditLogListTitle..." -ForegroundColor Cyan
New-PnPList -Title $auditLogListTitle -Template GenericList -Url $auditLogListName

# 4.6.a Vytvoření sloupců
Add-PnPField -List $auditLogListTitle -DisplayName "Detail" -InternalName "Detail" -Type Note -AddToDefaultView
Add-PnPField -List $auditLogListTitle -DisplayName "Člen ID" -InternalName "ClenID" -Type Number -AddToDefaultView
Add-PnPField -List $auditLogListTitle -DisplayName "Vydaná permanentka ID" -InternalName "VydanaPermanentkaID" -Type Number -AddToDefaultView
Add-PnPField -List $auditLogListTitle -DisplayName "Návštěva ID" -InternalName "NavstevaID" -Type Number -AddToDefaultView
Add-PnPField -List $auditLogListTitle -DisplayName "Závada ID" -InternalName "ZavadaID" -Type Number -AddToDefaultView

Write-Host "Seznam $auditLogListTitle a jeho pole byly úspěšně vytvořeny!" -ForegroundColor Green


# =========================================================================
# 4.7 Seznam - "Závady"
# =========================================================================
$zavadaListTitle = "Závada"
$zavadaListName = "Zavada"

Write-Host "Vytvářím seznam $zavadaListTitle..." -ForegroundColor Cyan
New-PnPList -Title $zavadaListTitle -Template GenericList -Url $zavadaListName

# 4.7.a Vytvoření specifických sloupců
Add-PnPField -List $zavadaListTitle -DisplayName "Popis závady" -InternalName "PopisZavady" -Type Note -AddToDefaultView
Add-PnPField -List $zavadaListTitle -DisplayName "Datum vzniku závady" -InternalName "DatumVzniku" -Type DateTime -AddToDefaultView
Set-PnPField -List $zavadaListTitle -Identity "DatumVzniku" -Values @{DisplayFormat="DateOnly"}
Add-PnPField -List $zavadaListTitle -DisplayName "Priorita" -InternalName "Priorita" -Type Choice -Choices "Nízká", "Střední", "Vysoká" -AddToDefaultView
Add-PnPField -List $zavadaListTitle -DisplayName "Fotografie závady" -InternalName "FotografieZavady" -Type Thumbnail -AddToDefaultView
Add-PnPField -List $zavadaListTitle -DisplayName "Stav" -InternalName "Stav" -Type Choice -Choices "Nový", "V řešení", "Vyřešeno" -AddToDefaultView

Write-Host "Seznam $zavadaListTitle a jeho pole byly úspěšně vytvořeny!" -ForegroundColor Green
# 1. Připojení k SharePoint webu
$siteUrl = "https://legranc.sharepoint.com/sites/UPCEGymManagement"
Connect-PnPOnline -Url $siteUrl -UseWebLogin

# Název seznamu
$listName = "Členství"
$pocetPolozek = 20

# 2. Zásobníky pro generování náhodných českých dat
$jmena = @("Jan", "Petr", "Martin", "Tomáš", "Jakub", "Jiří", "Michal", "Ondřej", "Lukáš", "David", "Kateřina", "Lucie", "Tereza", "Anna", "Veronika", "Monika", "Eliška", "Adéla")
$prijmeni = @("Novák", "Svoboda", "Novotný", "Dvořák", "Černý", "Procházka", "Kučera", "Veselý", "Krejčí", "Horák", "Nováková", "Svobodová", "Novotná", "Dvořáková", "Černá", "Procházková")
$poznamky = @("Aktivní sportovec", "Preferuje ranní tréninky", "Student FES", "Bezproblémový člen", "Ztracená karta v minulosti", "Nováček, ukázat stroje", "")

Write-Host "Zahajuji generování $pocetPolozek dummy členů do seznamu $listName..." -ForegroundColor Cyan

# 3. Smyčka pro generování dat
for ($i = 1; $i -le $pocetPolozek; $i++) {
    # Výběr náhodných prvků
    $jmeno = Get-Random -InputObject $jmena
    $prijmeniVybrane = Get-Random -InputObject $prijmeni
    $celeJmeno = "$jmeno $prijmeniVybrane"
    
    # Odstranění diakritiky pro NetID a Email
    $jmenoBezDiakritiky = [System.Text.Encoding]::ASCII.GetString([System.Text.Encoding]::GetEncoding("Cyrillic").GetBytes($jmeno)) -replace '[^a-zA-Z]', ''
    $prijmeniBezDiakritiky = [System.Text.Encoding]::ASCII.GetString([System.Text.Encoding]::GetEncoding("Cyrillic").GetBytes($prijmeniVybrane)) -replace '[^a-zA-Z]', ''
    
    # Generování specifických hodnot
    $netID = ("st" + (Get-Random -Minimum 10000 -Maximum 99999)) # Klasické studentské NetID stXXXXX
    $email = "$($jmenoBezDiakritiky.ToLower()).$($prijmeniBezDiakritiky.ToLower())@student.upce.cz"
    $telefon = "+420" + (Get-Random -Minimum 602000000 -Maximum 777999999)
    $nfcID = "NFC-" + (Get-Random -Minimum 10000000 -Maximum 99999999)
    
    $poznamka = Get-Random -InputObject $poznamky

    # Hashtable s vnitřními názvy sloupců (InternalName) a hodnotami
    $itemValues = @{
        "Title" = $celeJmeno
        "NetID" = $netID
        "Jmeno" = $jmeno
        "Prijmeni" = $prijmeniVybrane
        "Email" = $email
        "Telefon" = $telefon
        "NfcID" = $nfcID
        "Poznamka" = $poznamka
    }

    # Zápis do SharePointu
    try {
        Add-PnPListItem -List $listName -Values $itemValues | Out-Null
        Write-Host "[$i/$pocetPolozek] Přidán člen: $celeJmeno ($netID)" -ForegroundColor Green
    }
    catch {
        Write-Error "Chyba při přidávání člena $celeJmeno : $_"
    }
}

Write-Host "Hotovo! Všechna data byla úspěšně nahrána do SharePointu." -ForegroundColor Green


# =========================================================================
# Generování Permanentek
# =========================================================================
$listPermanentky = "Vydaná permanentka"
$listClenove = "Členství"

# Načteme si existující členy ze seznamu, abychom měli reálná ID a NetID pro vazby
Write-Host "Načítám členy ze seznamu $listClenove pro vytvoření vazeb..." -ForegroundColor Cyan
$clenove = Get-PnPListItem -List $listClenove -Fields "Id", "NetID", "Title"

if ($clenove.Count -eq 0) {
    Write-Error "V seznamu 'Clen' nebyly nalezeny žádné záznamy. Nejdříve prosím spusť skript na vygenerování členů!"
    return
}

# 2. Definice číselníků typů permanentek podle tvého zadání
$typyPermanentek = @(
    @{ Id = 1; Typ = "Jednorázový"; Vstupy = 1; Cena = 100 },
    @{ Id = 2; Typ = "10 vstupů"; Vstupy = 10; Cena = 800 },
    @{ Id = 3; Typ = "20 vstupů"; Vstupy = 20; Cena = 1200 }
)

$platebniMetody = @("Hotovost", "Bankovní převod", "Karta")
$stavyOstatni = @("Blokovaná", "Neaktivní")

Write-Host "Zahajuji generování permanentek..." -ForegroundColor Cyan

# Projedeme členy a každému přiřadíme 1 až 2 permanentky, ať máme hezká testovací data
$pocetVygenerovanych = 0

foreach ($clen in $clenove) {
    # Náhodně určíme, kolik permanentek tento člen dostane (např. 1 nebo 2)
    $pocetPermicNaClena = Get-Random -Minimum 1 -Maximum 3
    
    for ($j = 0; $j -lt $pocetPermicNaClena; $j++) {
        # Výběr náhodného typu permanentky
        $typ = Get-Random -InputObject $typyPermanentek
        
        # Generování logických dat pro platnost (v rozmezí +- 3 měsíce od teď)
        $dnyPosunOdDneska = Get-Random -Minimum -90 -Maximum 60
        $platnostOd = (Get-Date).AddDays($dnyPosunOdDneska)
        
        # Permanentka platí standardně 3 měsíce (90 dní) od vystavení
        $platnostDo = $platnostOd.AddDays(90)
        
        # Automatické určení stavu na základě datumu
        $stav = "Aktivní"
        if ($platnostDo -lt (Get-Date)) {
            $stav = "Expirovaná"
        } else {
            # Malá šance (10%), že permanentka bude blokovaná nebo neaktivní i přes platné datum
            if ((Get-Random -Minimum 1 -Maximum 11) -eq 1) {
                $stav = Get-Random -InputObject $stavyOstatni
            }
        }
        
        # Tento kousek kódu jsem opravil, aby odpovídal původní logice
        $pocetMoznychVstupu = $typ.Vstupy
        if ($stav -eq "Expirovaná") {
            $pocetZbyvajicichVstupu = Get-Random -Minimum 0 -Maximum 3
            if ($pocetZbyvajicichVstupu -gt $pocetMoznychVstupu) { $pocetZbyvajicichVstupu = 0 }
        } else {
            $pocetZbyvajicichVstupu = Get-Random -Minimum 0 -Maximum ($pocetMoznychVstupu + 1)
        }
        
        $jeZaplacena = if ($stav -eq "Neaktivní") { $false } else { $true }
        $platebniMetoda = Get-Random -InputObject $platebniMetody
        
        $title = "$($typ.Typ) - $($clen["NetID"])"

        # Složení hodnot pro SharePoint řádek
        $itemValues = @{
            "Title"                   = $clen["Title"]
            "PlatnostOd"              = $platnostOd
            "PlatnostDo"              = $platnostDo
            "JeZaplacena"             = $jeZaplacena
            "Stav"                    = $stav
            "PlatebniMetoda"          = $platebniMetoda
            "NetID"                   = $clen["NetID"]
            "ClenID"                  = $clen.Id
            "TypPermanentkyID"        = $typ.Id
            "PocetMoznychVstupu"      = $pocetMoznychVstupu
            "PocetZbyvajicichVstupu"  = $pocetZbyvajicichVstupu
            "TypPermanentky"          = $typ.Typ
            "Cena"                    = $typ.Cena
        }

        try {
            Add-PnPListItem -List $listPermanentky -Values $itemValues | Out-Null
            $pocetVygenerovanych++
            Write-Host "Vytvořena permanentka pro $($clen["NetID"]): $title (Stav: $stav)" -ForegroundColor Green
        }
        catch {
            Write-Error "Selhalo vytvoření permanentky pro člena s ID $($clen.Id): $_"
        }
    }
}

Write-Host "Hotovo! Celkem úspěšně vygenerováno $pocetVygenerovanych permanentek." -ForegroundColor Green

# =========================================================================
# Generování Návštěv
# =========================================================================
$navstevaListTitle = "Návštěva"

Write-Host "Načítám vygenerované permanentky ze seznamu '$listPermanentky'..." -ForegroundColor Cyan
# Načteme si permanentky, abychom znali jejich ID a ID člena, kterému patří
$permanentky = Get-PnPListItem -List $listPermanentky -Fields "Id", "ClenID", "Title"

if ($permanentky.Count -eq 0) {
    Write-Error "Nebyly nalezeny žádné permanentky. Nelze generovat návštěvy!"
    return
}

Write-Host "Zahajuji generování náhodných návštěv..." -ForegroundColor Cyan
$pocetNavstev = 0

# Projdeme permanentky a pro každou vytvoříme několik náhodných návštěv (např. 1 až 4)
foreach ($permice in $permanentky) {
    $pocetVstupuNaPermici = Get-Random -Minimum 1 -Maximum 5
    
    # Získání celého jména člena přímo z permanentky
    $celeJmenoClena = $permice["Title"]
    $clenId = $permice["ClenID"]
    $permiceId = $permice.Id

    for ($k = 0; $k -lt $pocetVstupuNaPermici; $k++) {
        
        # Generování náhodného času příchodu (v rozmezí posledních 30 dní)
        $dnyZpet = Get-Random -Minimum 0 -Maximum 30
        $hodinaPrichodu = Get-Random -Minimum 14 -Maximum 21 # Gym je otevřený od 14 do 21
        $minutaPrichodu = Get-Random -Minimum 0 -Maximum 60
        
        $objektPrichodu = (Get-Date).AddDays(-$dnyZpet)
        $objektPrichodu = Get-Date -Date $objektPrichodu -Hour $hodinaPrichodu -Minute $minutaPrichodu -Second 0
        
        # Délka tréninku náhodně mezi 1 až 3 hodinami (60 až 180 minut)
        $delkaTreninkuMinuty = Get-Random -Minimum 60 -Maximum 181
        $objektOdchodu = $objektPrichodu.AddMinutes($delkaTreninkuMinuty)

        # Převedení DateTime objektů na ISO textový formát pro SharePoint
        $casPrichoduISO = $objektPrichodu.ToString("yyyy-MM-ddTHH:mm:ssZ")
        $casOdchoduISO = $objektOdchodu.ToString("yyyy-MM-ddTHH:mm:ssZ")

        # Příprava hodnot pro zápis do seznamu Návštěva
        $navstevaValues = @{
            "Title"                = $celeJmenoClena        
            "CasPrichodu"          = $casPrichoduISO    # Posíláme jako ISO text
            "CasOdchodu"           = $casOdchoduISO     # Posíláme jako ISO text
            "VydanaPermanentkaID"  = [int]$permiceId
            "ClenID"               = [int]$clenId
        }

        try {
            Add-PnPListItem -List $navstevaListTitle -Values $navstevaValues | Out-Null
            $pocetNavstev++
            Write-Host "[$pocetNavstev] Zapsána návštěva: $celeJmenoClena (Příchod: $objektPrichodu)" -ForegroundColor Green
        }
        catch {
            Write-Error "Chyba při zápisu návštěvy pro člena ID $clenId"
        }

        Start-Sleep -Milliseconds 50
    }
}

Write-Host "Vše hotovo! Celkem úspěšně vygenerováno $pocetNavstev návštěv." -ForegroundColor Green
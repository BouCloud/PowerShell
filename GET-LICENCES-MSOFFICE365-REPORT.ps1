<#
.SYNOPSIS
    Rapport des licences Microsoft 365 via Microsoft Graph et export Excel.

.DESCRIPTION
    - Vérifie / installe les modules suivants (Scope = CurrentUser) :
        * ImportExcel
        * Microsoft.Graph (pour disposer de Microsoft.Graph.Authentication)
    - Télécharge le CSV officiel Microsoft permettant de faire le mapping SKU -> Nom de produit.
    - Récupère les données de licences via Microsoft Graph :
        * /beta/directory/subscriptions
        * /v1.0/subscribedSkus
    - Calcule pour chaque SKU :
        * Total / Consommé / Disponible
        * Pourcentage utilisé
        * Barre de progression visuelle (texte)
    - Affiche le rapport dans la console.
    - Exporte le rapport dans un fichier Excel :
        C:\temp\Licenses_Report_yyyy_MM_dd.xlsx

    Remarques :
        - PowerShell ISE n’est pas supporté pour Microsoft.Graph.
          Ce script ne s’exécute pas dans ISE.
        - Merci de relire et adapter le code à votre contexte avant utilisation en production.

.AUTEUR
    Christian BOUSQUET
#>

# --------------------------------------------------------------------
# Paramétrage de base
# --------------------------------------------------------------------

# Mode d’authentification :
#   - "Interactive" : fenêtre de login (recommandé en PowerShell 5.1)
#   - "DeviceCode"  : code appareil (utile sur serveur / RDP si l’interactif pose problème)
#   ! "Auto" est volontairement désactivé à cause de bugs DeviceCodeCredential en PS 5.1.
$AuthMode = "Interactive"

# Petite fonction de pause uniquement en cas d’erreur, et uniquement si la session est interactive
function Pause-IfInteractive {
    # UserInteractive = True => session avec UI (pas tâche planifiée en arrière-plan)
    if ([Environment]::UserInteractive -and $Host.Name -eq 'ConsoleHost') {
        try {
            Write-Host ""
            Read-Host "Une erreur est survenue. Appuyez sur Entrée pour fermer cette fenêtre"
        }
        catch {
            # Si jamais Read-Host pose problème en non interactif, on ignore.
        }
    }
}


#region --- [0/5] Vérifications préalables (ISE, version PowerShell) ---

# Blocage si PowerShell ISE
if ($psISE) {
    Write-Warning "Ce script ne s'exécute pas dans PowerShell ISE."
    Write-Warning "Microsoft Graph n'est pas officiellement supporté dans ISE."
    Write-Host   ""
    Write-Host "Veuillez lancer ce script depuis :" -ForegroundColor Yellow
    Write-Host "  - une console PowerShell classique (Windows PowerShell)" -ForegroundColor Yellow
    Write-Host "  - ou PowerShell 7 (pwsh)" -ForegroundColor Yellow
    Write-Host   ""
    Read-Host "Appuyez sur Entrée pour quitter"
    return
}

# Version minimale recommandée
$minVersion = [version]'5.1'
if ($PSVersionTable.PSVersion -lt $minVersion) {
    Write-Warning "Version PowerShell détectée : $($PSVersionTable.PSVersion). Il est recommandé d'utiliser au moins PowerShell 5.1 ou PowerShell 7."
}

Write-Host "[1/5] Environnement PowerShell : $($PSVersionTable.PSVersion) - Edition : $($PSVersionTable.PSEdition)" -ForegroundColor Cyan

#endregion


#region --- [1/5] Installation / chargement des modules ---

Write-Host "  - Les modules seront installés dans le profil utilisateur (Scope = CurrentUser, pas besoin d'être administrateur)." -ForegroundColor DarkGray

# Provider NuGet pour PowerShellGet (évite les questions interactives)
try {
    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
        Write-Host "  - Provider NuGet absent. Installation (Scope = CurrentUser)..." -ForegroundColor Yellow
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction Stop
        Write-Host "  - Provider NuGet installé." -ForegroundColor Green
    }
}
catch {
    Write-Warning "Impossible d'installer le provider NuGet automatiquement. Les installations de modules peuvent demander une confirmation interactive."
}

# Liste des modules requis
$requiredModules = @(
    'ImportExcel',
    'Microsoft.Graph'
)

$moduleCount = $requiredModules.Count
$index       = 0

foreach ($moduleName in $requiredModules) {
    $index++

    if (-not (Get-Module -ListAvailable -Name $moduleName)) {

        $percent = $index / $moduleCount * 100
        Write-Progress -Activity "Installation des modules requis" `
                       -Status "Installation de '$moduleName' (Scope = CurrentUser)..." `
                       -PercentComplete $percent

        # Message spécifique pour Microsoft.Graph
        if ($moduleName -eq 'Microsoft.Graph') {
            Write-Host ""
            Write-Host "  ⚠ Le module 'Microsoft.Graph' est volumineux (plusieurs sous-modules, milliers de commandes)." -ForegroundColor Yellow
            Write-Host "  ⚠ Son installation peut prendre un certain temps, surtout lors du premier téléchargement." -ForegroundColor Yellow
            Write-Host "  ⚠ C'est normal, merci de patienter jusqu'à la fin de l'installation..." -ForegroundColor Yellow
            Write-Host ""
        }

        Write-Host "  - Le module '$moduleName' n'est pas installé. Installation en cours (Scope = CurrentUser)..." -ForegroundColor Yellow
        try {
            Install-Module -Name $moduleName -Scope CurrentUser -Force -Confirm:$false -ErrorAction Stop
            Write-Host "  - Module '$moduleName' installé avec succès." -ForegroundColor Green
        }
        catch {
            Write-Error "Échec de l'installation du module '$moduleName'. Détails : $_"
            Write-Progress -Activity "Installation des modules requis" -Completed
            Pause-IfInteractive
            return
        }
    }
    else {
        Write-Host "  - Module '$moduleName' déjà présent." -ForegroundColor Green
    }
}

Write-Progress -Activity "Installation des modules requis" -Completed

# Import des modules nécessaires à l'exécution
try {
    Write-Host "  - Import du module ImportExcel..." -ForegroundColor Yellow
    Import-Module ImportExcel -ErrorAction Stop
    Write-Host "  - Module ImportExcel importé." -ForegroundColor Green

    Write-Host "  - Import du module Microsoft.Graph.Authentication..." -ForegroundColor Yellow
    Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue
    Write-Host "  - Module Microsoft.Graph.Authentication disponible." -ForegroundColor Green
}
catch {
    Write-Error "Erreur lors de l'import des modules : $_"
    Pause-IfInteractive
    return
}

#endregion


#region --- [2/5] Préparation du dossier et du CSV Microsoft ---

Write-Host "[2/5] Préparation du dossier d'export et du CSV Microsoft..." -ForegroundColor Cyan

$exportFolder = "C:\temp"
if (-not (Test-Path -Path $exportFolder)) {
    New-Item -ItemType Directory -Path $exportFolder | Out-Null
    Write-Host "  - Dossier $exportFolder créé." -ForegroundColor Green
}
else {
    Write-Host "  - Dossier $exportFolder déjà présent." -ForegroundColor Green
}

$csvSource = Join-Path -Path $exportFolder -ChildPath "Product names and service plan identifiers for licensing.csv"

if (-not (Test-Path $csvSource)) {
    try {
        Write-Host "  - Téléchargement du CSV de référence Microsoft..." -ForegroundColor Yellow
        Invoke-WebRequest -Uri "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv" -OutFile $csvSource -ErrorAction Stop
        Write-Host "  - CSV de référence téléchargé : $csvSource" -ForegroundColor Green
    }
    catch {
        Write-Error "Échec du téléchargement du CSV : $_"
        Pause-IfInteractive
        return
    }
}
else {
    Write-Host "  - CSV de référence déjà présent : $csvSource" -ForegroundColor Green
}

# Mapping SKU -> Nom de produit
$productMap = @{}
try {
    Write-Host "  - Chargement du mapping produits depuis le CSV..." -ForegroundColor Yellow
    Import-Csv $csvSource -Encoding UTF8 | ForEach-Object {
        $sku = $_.'String_Id'.Trim()
        if (-not [string]::IsNullOrWhiteSpace($sku)) {
            $productMap[$sku] = $_.'Product_Display_Name'.Trim()
        }
    }
    Write-Host "  - Mapping des produits chargé." -ForegroundColor Green
}
catch {
    Write-Error "Erreur de traitement du CSV : $_"
    Pause-IfInteractive
    return
}

#endregion


#region --- [3/5] Connexion à Microsoft Graph ---

Write-Host "[3/5] Connexion à Microsoft Graph..." -ForegroundColor Cyan

# Profil beta pour /directory/subscriptions
try {
    Select-MgProfile -Name "beta" -ErrorAction SilentlyContinue
} catch { }

# Désactivation de WAM pour éviter certains problèmes d’UI
try {
    Set-MgGraphOption -DisableLoginByWAM $true -ErrorAction SilentlyContinue
} catch { }

$scopes = @(
    "Directory.Read.All",
    "Organization.Read.All",
    "User.Read.All"
)

function Show-MgContextInfo {
    $ctx = Get-MgContext
    if ($ctx) {
        Write-Host ("  - Connecté à Microsoft Graph. TenantId : {0}, Compte ou Id : {1}" -f $ctx.TenantId, $ctx.Account) -ForegroundColor Green
        Write-Host ("  - Mode d'authentification : {0}" -f $AuthMode) -ForegroundColor Green
    }
    else {
        Write-Warning "  - Contexte Microsoft Graph introuvable après connexion."
    }
}

function Connect-M365Graph {
    param(
        [ValidateSet("Interactive","DeviceCode")]
        [string]$Mode
    )

    switch ($Mode) {
        "Interactive" {
            Write-Host "  - Connexion interactive (fenêtre de login)..." -ForegroundColor Yellow
            Connect-MgGraph -Scopes $scopes -ErrorAction Stop
        }
        "DeviceCode" {
            Write-Host "  - Connexion via code appareil (Device Code)..." -ForegroundColor Yellow
            Connect-MgGraph -Scopes $scopes -UseDeviceCode -ErrorAction Stop
        }
    }
}

try {
    switch ($AuthMode) {

        "Interactive" {
            Connect-M365Graph -Mode "Interactive"
            Show-MgContextInfo
        }

        "DeviceCode" {
            Connect-M365Graph -Mode "DeviceCode"
            Show-MgContextInfo
        }

        default {
            Write-Warning "Valeur AuthMode '$AuthMode' invalide. Utilisation du mode 'Interactive'."
            Connect-M365Graph -Mode "Interactive"
            Show-MgContextInfo
        }
    }
}
catch {
    Write-Error "Échec de la connexion à Microsoft Graph : $($_.Exception.Message)"
    Pause-IfInteractive
    return
}

#endregion


#region --- [4/5] Récupération des données de licences ---

Write-Host "[4/5] Récupération des données de licences..." -ForegroundColor Cyan

try {
    Write-Host "  - Récupération des abonnements (/beta/directory/subscriptions)..." -ForegroundColor Yellow
    $subscriptions = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/directory/subscriptions" -ErrorAction Stop

    Write-Host "  - Récupération des SKU (/v1.0/subscribedSkus)..." -ForegroundColor Yellow
    $skuDetails    = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/subscribedSkus" -ErrorAction Stop
}
catch {
    Write-Error "Erreur lors de l'appel à Microsoft Graph : $($_.Exception.Message)"
    Pause-IfInteractive
    return
}

if (-not $subscriptions.value) {
    Write-Warning "  - Aucun abonnement retourné par /beta/directory/subscriptions."
}
else {
    Write-Host ("  - {0} abonnement(s) récupéré(s)." -f $subscriptions.value.Count) -ForegroundColor Green
}

if (-not $skuDetails.value) {
    Write-Warning "  - Aucun SKU retourné par /v1.0/subscribedSkus."
}
else {
    Write-Host ("  - {0} SKU récupéré(s)." -f $skuDetails.value.Count) -ForegroundColor Green
}

# Table de correspondance SKU -> consommé / activé
$skuHash = @{}
$skuDetails.value | ForEach-Object {
    if ($_.skuPartNumber) {
        $skuHash[$_.skuPartNumber] = @{
            Consumed = $_.consumedUnits
            Enabled  = $_.prepaidUnits.enabled
        }
    }
}

#endregion


#region --- [5/5] Traitement, affichage et export Excel ---

Write-Host "[5/5] Traitement des données et export..." -ForegroundColor Cyan

$result = @()

foreach ($sub in $subscriptions.value) {
    $skuPartNumber = $sub.skuPartNumber

    if ([string]::IsNullOrWhiteSpace($skuPartNumber)) {
        Write-Warning "  - Entrée subscription sans skuPartNumber, ignorée."
        continue
    }

    if ($skuHash.ContainsKey($skuPartNumber)) {
        $total     = [int]$sub.totalLicenses
        $consumed  = [int]$skuHash[$skuPartNumber].Consumed
        $available = $total - $consumed
        if ($available -lt 0) { $available = 0 }

        $usedPercent = if ($total -gt 0) {
            [math]::Min(100, [math]::Round(($consumed / $total) * 100, 1))
        } else { 0 }

        $filled      = [math]::Max(0, [math]::Min(20, [math]::Round($usedPercent / 5)))
        $empty       = 20 - $filled
        $progressBar = "[" + ("█" * $filled) + ("░" * $empty) + "]"

        $nomProduit = if ($productMap.ContainsKey($skuPartNumber)) {
            $productMap[$skuPartNumber]
        } else {
            "INCONNU ($skuPartNumber)"
        }

        $dateCreatedRaw   = $sub.createdDateTime
        $nextLifecycleRaw = $sub.nextLifecycleDateTime

        $dateCreated   = if ($dateCreatedRaw)   { $dateCreatedRaw -replace "T|Z"," " } else { $null }
        $nextLifecycle = if ($nextLifecycleRaw) { $nextLifecycleRaw -replace "T|Z"," " } else { $null }

        $result += [PSCustomObject]@{
            Produit           = $nomProduit
            SkuPartNumber     = $skuPartNumber
            TotalLicenses     = $total
            ConsumedLicenses  = $consumed
            AvailableLicenses = $available
            '% Utilisé'       = "$usedPercent%"
            Progression       = $progressBar
            IsTrial           = $sub.isTrial
            Status            = $sub.status
            DateCreated       = $dateCreated
            NextLifeCycleDate = $nextLifecycle
        }
    }
    else {
        Write-Warning "  - Aucune correspondance trouvée dans /subscribedSkus pour le SKU $skuPartNumber"
    }
}

if (-not $result) {
    Write-Warning "Aucune donnée à afficher/exporter."
    return
}

Write-Host "  - Affichage du rapport dans la console..." -ForegroundColor Yellow

$result |
    Sort-Object { $_.'% Utilisé'.Trim('%') -as [double] } -Descending |
    Format-Table -AutoSize -Property Produit,SkuPartNumber,TotalLicenses,ConsumedLicenses,AvailableLicenses,'% Utilisé',Progression,IsTrial,Status,DateCreated,NextLifeCycleDate

$excelPath = Join-Path -Path $exportFolder -ChildPath ("Licenses_Report_{0}.xlsx" -f (Get-Date -Format 'yyyy_MM_dd'))

Write-Host "  - Export du rapport vers Excel : $excelPath" -ForegroundColor Yellow
$result | Export-Excel -Path $excelPath -AutoSize -WorksheetName "Licences365" -TableName "TableLicences365" -TableStyle "Medium2"

Write-Host "  - Rapport Excel exporté vers : $excelPath" -ForegroundColor Green

#endregion
pause

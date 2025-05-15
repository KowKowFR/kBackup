##[Ps1 To Exe]
##
##Kd3HDZOFADWE8uO1
##Nc3NCtDXTlaDjofG5iZk2WjhUG0lesyV+YKS7qCbys/QnyrOR5YbSFBkqiXzA0TzUPEdNQ==
##Kd3HFJGZHWLWoLaVvnQnhQ==
##LM/RF4eFHHGZ7/K1
##K8rLFtDXTiW5
##OsHQCZGeTiiZ4NI=
##OcrLFtDXTiW5
##LM/BD5WYTiiZ4tI=
##McvWDJ+OTiiZ4tI=
##OMvOC56PFnzN8u+Vs1Q=
##M9jHFoeYB2Hc8u+Vs1Q=
##PdrWFpmIG2HcofKIo2QX
##OMfRFJyLFzWE8uK1
##KsfMAp/KUzWJ0g==
##OsfOAYaPHGbQvbyVvnQX
##LNzNAIWJGmPcoKHc7Do3uAuO
##LNzNAIWJGnvYv7eVvnQX
##M9zLA5mED3nfu77Q7TV64AuzAgg=
##NcDWAYKED3nfu77Q7TV64AuzAgg=
##OMvRB4KDHmHQvbyVvnQX
##P8HPFJGEFzWE8tI=
##KNzDAJWHD2fS8u+Vgw==
##P8HSHYKDCX3N8u+Vgw==
##LNzLEpGeC3fMu77Ro2k3hQ==
##L97HB5mLAnfMu77Ro2k3hQ==
##P8HPCZWEGmaZ7/K1
##L8/UAdDXTlaDjofG5iZk2WjhUG0lesyV+YKS7qCbys/QnyraXJcRR0Bkqgz9F1L9eOgHR/BVgN4eWQ5kKuoOgg==
##Kc/BRM3KXxU=
##
##
##fd6a9f26a06ea3bc99616d4851b372ba

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

function Install-RequiredModule {
    param (
        [string]$ModuleName
    )
    
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "Le module $ModuleName n'est pas installé. Installation en cours..." -ForegroundColor Yellow
        
        if (Test-Admin) {
            try {
                Install-Module -Name $ModuleName -Force -AllowClobber -Scope AllUsers
                Write-Host "Le module $ModuleName a été installé avec succès." -ForegroundColor Green
                Connect-AzureAD
            }
            catch {
                Write-Host "Erreur lors de l'installation du module $ModuleName : $_" -ForegroundColor Red
                return $false
            }
        }
        else {
            # Relancer le script en tant qu'administrateur
            Write-Host "Élévation des privilèges nécessaire pour installer le module $ModuleName." -ForegroundColor Yellow
            
            $scriptPath = $MyInvocation.MyCommand.Path
            $arguments = "-NoExit -ExecutionPolicy Bypass -File `"$scriptPath`""
            
            try {
                Start-Process powershell.exe -Verb RunAs -ArgumentList $arguments
                Write-Host "Le script a été relancé avec des privilèges administrateur. Veuillez fermer cette fenêtre." -ForegroundColor Green
                exit
            }
            catch {
                Write-Host "Impossible d'élever les privilèges. Veuillez exécuter ce script en tant qu'administrateur." -ForegroundColor Red
                return $false
            }
        }
    }
    else {
        Write-Host "Le module $ModuleName est déjà installé." -ForegroundColor Green
        Connect-AzureAD
    }
    
    return $true
}

# Vérifier et installer le module AzureAD
$moduleInstalled = Install-RequiredModule -ModuleName "AzureAD"
if (-not $moduleInstalled) {
    Write-Host "Le module AzureAD est requis pour exécuter ce script. Veuillez l'installer manuellement." -ForegroundColor Red
    Read-Host "Appuyez sur Entrée pour quitter"
    exit
}

# Importation du module AzureAD
try {
    Import-Module AzureAD -ErrorAction Stop
    Write-Host "Module AzureAD importé avec succès." -ForegroundColor Green
}
catch {
    Write-Host "Erreur lors de l'importation du module AzureAD : $_" -ForegroundColor Red
    Read-Host "Appuyez sur Entrée pour quitter"
    exit
}

# Importation des modules necessaires
# Si le module AzureAD n'est pas installe, vous devrez l'installer avec :
# Install-Module -Name AzureAD -Force -AllowClobber

# Chargement des assemblies pour l'interface graphique
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$logoBase64 = "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"

# Definition des couleurs pour une interface moderne
$primaryColor = [System.Drawing.Color]::FromArgb(0, 120, 212)     # Bleu Microsoft
$accentColor = [System.Drawing.Color]::FromArgb(0, 200, 83)       # Vert pour succes
$errorColor = [System.Drawing.Color]::FromArgb(209, 52, 56)       # Rouge pour erreur
$textColor = [System.Drawing.Color]::FromArgb(50, 49, 48)         # Gris fonce pour texte
$bgColor = [System.Drawing.Color]::FromArgb(250, 250, 250)        # Fond legerement grise
$cardBgColor = [System.Drawing.Color]::White                      # Fond blanc pour les cartes

# Afficher un popup a l'ouverture pour indiquer l'auteur du script
$authorPopup = New-Object System.Windows.Forms.Form
$authorPopup.Text = "Information"
$authorPopup.Size = New-Object System.Drawing.Size(450, 250)
$authorPopup.StartPosition = "CenterScreen"
$authorPopup.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$authorPopup.MaximizeBox = $false
$authorPopup.MinimizeBox = $false
$authorPopup.BackColor = $bgColor
$authorPopup.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# Ajouter un logo dans le popup (optionnel)
$popupLogoBox = New-Object System.Windows.Forms.PictureBox
$popupLogoBox.Location = New-Object System.Drawing.Point(20, 20)
$popupLogoBox.Size = New-Object System.Drawing.Size(40, 40)
$popupLogoBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom

function ConvertFrom-Base64ToImage {
    param (
        [string]$Base64String
    )
    
    try {
        $imageBytes = [Convert]::FromBase64String($Base64String)
        $memoryStream = New-Object System.IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
        $image = [System.Drawing.Image]::FromStream($memoryStream)
        return $image
    }
    catch {
        Write-Host "Erreur lors de la conversion de l'image: $($_.Exception.Message)"
        return $null
    }
}

$logoImage = ConvertFrom-Base64ToImage -Base64String $logoBase64
if ($logoImage -ne $null) {
    $popupLogoBox.Image = $logoImage
}
else {
    # Créer un logo par défaut si la conversion échoue
    try {
        $popupIcon = New-Object System.Drawing.Bitmap(40, 40)
        $graphics = [System.Drawing.Graphics]::FromImage($popupIcon)
        $graphics.Clear([System.Drawing.Color]::White)
        $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(0, 120, 212), 2)
        $graphics.DrawRectangle($pen, 5, 5, 30, 30)
        $graphics.DrawLine($pen, 5, 20, 35, 20)
        $graphics.DrawLine($pen, 20, 5, 20, 35)
        $graphics.Dispose()
        $popupLogoBox.Image = $popupIcon
    }
    catch {
        # En cas d'erreur, ne pas afficher de logo
    }
}
$authorPopup.Controls.Add($popupLogoBox)

# Ajouter le message d'information
$infoLabel = New-Object System.Windows.Forms.Label
$infoLabel.Location = New-Object System.Drawing.Point(70, 20)
$infoLabel.Size = New-Object System.Drawing.Size(400, 60)
$infoLabel.Text = "Ce script a ete cree par Corentin, Xavier, Samy et Younes.`n`nInterface de gestion Azure Active Directory."
$infoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$infoLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$authorPopup.Controls.Add($infoLabel)

# Ajouter les boutons LinkedIn
$linkedinLabel = New-Object System.Windows.Forms.Label
$linkedinLabel.Location = New-Object System.Drawing.Point(70, 90)
$linkedinLabel.Size = New-Object System.Drawing.Size(300, 20)
$linkedinLabel.Text = "Retrouvez-nous sur LinkedIn :"
$linkedinLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$linkedinLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$authorPopup.Controls.Add($linkedinLabel)

# Bouton LinkedIn pour Corentin
$linkedinCorentin = New-Object System.Windows.Forms.Button
$linkedinCorentin.Location = New-Object System.Drawing.Point(70, 115)
$linkedinCorentin.Size = New-Object System.Drawing.Size(70, 25)
$linkedinCorentin.Text = "Corentin"
$linkedinCorentin.BackColor = [System.Drawing.Color]::FromArgb(10, 102, 194)
$linkedinCorentin.ForeColor = [System.Drawing.Color]::White
$linkedinCorentin.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$linkedinCorentin.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Regular)
$linkedinCorentin.Add_Click({ Start-Process "https://www.linkedin.com/in/corentin-tujague" })
$authorPopup.Controls.Add($linkedinCorentin)

# Bouton LinkedIn pour Xavier
$linkedinXavier = New-Object System.Windows.Forms.Button
$linkedinXavier.Location = New-Object System.Drawing.Point(145, 115)
$linkedinXavier.Size = New-Object System.Drawing.Size(70, 25)
$linkedinXavier.Text = "Xavier"
$linkedinXavier.BackColor = [System.Drawing.Color]::FromArgb(10, 102, 194)
$linkedinXavier.ForeColor = [System.Drawing.Color]::White
$linkedinXavier.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$linkedinXavier.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Regular)
$linkedinXavier.Add_Click({ Start-Process "https://www.linkedin.com/in/xavier-lunel/" })
$authorPopup.Controls.Add($linkedinXavier)

# Bouton LinkedIn pour Samy
$linkedinSamy = New-Object System.Windows.Forms.Button
$linkedinSamy.Location = New-Object System.Drawing.Point(220, 115)
$linkedinSamy.Size = New-Object System.Drawing.Size(70, 25)
$linkedinSamy.Text = "Samy"
$linkedinSamy.BackColor = [System.Drawing.Color]::FromArgb(10, 102, 194)
$linkedinSamy.ForeColor = [System.Drawing.Color]::White
$linkedinSamy.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$linkedinSamy.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Regular)
$linkedinSamy.Add_Click({ Start-Process "https://www.linkedin.com/in/samy-argoub-65b7bb24a/" })
$authorPopup.Controls.Add($linkedinSamy)

# Bouton LinkedIn pour Younes
$linkedinYounes = New-Object System.Windows.Forms.Button
$linkedinYounes.Location = New-Object System.Drawing.Point(295, 115)
$linkedinYounes.Size = New-Object System.Drawing.Size(70, 25)
$linkedinYounes.Text = "Younes"
$linkedinYounes.BackColor = [System.Drawing.Color]::FromArgb(10, 102, 194)
$linkedinYounes.ForeColor = [System.Drawing.Color]::White
$linkedinYounes.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$linkedinYounes.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Regular)
$linkedinYounes.Add_Click({ Start-Process "https://www.linkedin.com/in/younes-ounadjela/" })
$authorPopup.Controls.Add($linkedinYounes)

# Ajouter un bouton GitHub pour EasyFormer
$githubButton = New-Object System.Windows.Forms.Button
$githubButton.Location = New-Object System.Drawing.Point(70, 145)
$githubButton.Size = New-Object System.Drawing.Size(295, 30)
$githubButton.Text = "Visitez notre GitHub EasyFormer"
$githubButton.BackColor = [System.Drawing.Color]::FromArgb(36, 41, 46)  # Couleur GitHub
$githubButton.ForeColor = [System.Drawing.Color]::White
$githubButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$githubButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$githubButton.Add_Click({ Start-Process "https://github.com/easyformer" })
$authorPopup.Controls.Add($githubButton)

# Ajouter un bouton OK (déplacé plus bas pour faire de la place au bouton GitHub)
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(175, 180)  # Position ajustée
$okButton.Size = New-Object System.Drawing.Size(100, 30)
$okButton.Text = "OK"
$okButton.BackColor = $primaryColor
$okButton.ForeColor = [System.Drawing.Color]::White
$okButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$okButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
$okButton.Add_Click({ $authorPopup.Close() })
$authorPopup.Controls.Add($okButton)

# Afficher le popup
[void]$authorPopup.ShowDialog()

# Creation de la fenetre principale avec un design plus moderne
$form = New-Object System.Windows.Forms.Form
$form.Text = "Azure AD - Interface Graphique"
$form.Size = New-Object System.Drawing.Size(600, 850)  # Augmentation de la hauteur de 800 à 850 px
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.BackColor = $bgColor
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# Création d'un panel d'en-tête avec couleur d'accentuation
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock = [System.Windows.Forms.DockStyle]::Top
$headerPanel.Height = 70
$headerPanel.BackColor = $primaryColor
$form.Controls.Add($headerPanel)

# Chemin vers le logo Azure AD (à remplacer par le chemin de votre logo)
$logoBox = New-Object System.Windows.Forms.PictureBox
$logoBox.Location = New-Object System.Drawing.Point(20, 15)
$logoBox.Size = New-Object System.Drawing.Size(40, 40)
$logoBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom

# Utiliser la même image encodée en Base64 pour le logo principal
if ($logoImage -ne $null) {
    $logoBox.Image = $logoImage
}
else {
    # Créer un logo par défaut si nécessaire
    try {
        $azureIcon = New-Object System.Drawing.Bitmap(40, 40)
        $graphics = [System.Drawing.Graphics]::FromImage($azureIcon)
        $graphics.Clear([System.Drawing.Color]::White)
        $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(0, 120, 212), 2)
        $graphics.DrawRectangle($pen, 5, 5, 30, 30)
        $graphics.DrawLine($pen, 5, 20, 35, 20)
        $graphics.DrawLine($pen, 20, 5, 20, 35)
        $graphics.Dispose()
        $logoBox.Image = $azureIcon
    }
    catch {
        # En cas d'erreur, ne pas afficher de logo
        Write-Host "Impossible de créer un logo par défaut : $($_.Exception.Message)"
    }
}

$headerPanel.Controls.Add($logoBox)

# Création du titre dans l'en-tête (ajustement de la position pour faire place au logo)
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location = New-Object System.Drawing.Point(70, 20)
$titleLabel.Size = New-Object System.Drawing.Size(470, 30)
$titleLabel.Text = "Azure Active Directory Interface"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::White
$titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$headerPanel.Controls.Add($titleLabel)



# Création d'un conteneur principal pour le contenu
$mainContainer = New-Object System.Windows.Forms.Panel
$mainContainer.Location = New-Object System.Drawing.Point(0, 70)
$mainContainer.Size = New-Object System.Drawing.Size(600, 780)  # Augmentation de la taille de 730 à 780 px
$mainContainer.BackColor = $bgColor
$mainContainer.Padding = New-Object System.Windows.Forms.Padding(20, 20, 20, 20)
$mainContainer.AutoScroll = $true  # Activation du défilement automatique
$form.Controls.Add($mainContainer)

# Création d'une carte pour le bouton de connexion
$connectionCard = New-Object System.Windows.Forms.Panel
$connectionCard.Location = New-Object System.Drawing.Point(30, 20)
$connectionCard.Size = New-Object System.Drawing.Size(540, 70)
$connectionCard.BackColor = $cardBgColor
$connectionCard.BorderStyle = [System.Windows.Forms.BorderStyle]::None
# Ajouter une ombre (simulée avec une bordure)
$connectionCard.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$mainContainer.Controls.Add($connectionCard)

# Texte d'instruction
$instructionLabel = New-Object System.Windows.Forms.Label
$instructionLabel.Location = New-Object System.Drawing.Point(20, 10)
$instructionLabel.Size = New-Object System.Drawing.Size(300, 50)
$instructionLabel.Text = "Connectez-vous pour acceder a vos informations Azure AD"
$instructionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$instructionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$connectionCard.Controls.Add($instructionLabel)

# Création du bouton de connexion avec style moderne
$connectButton = New-Object System.Windows.Forms.Button
$connectButton.Location = New-Object System.Drawing.Point(350, 15)
$connectButton.Size = New-Object System.Drawing.Size(170, 40)
$connectButton.Text = "Se connecter"
$connectButton.BackColor = $primaryColor
$connectButton.ForeColor = [System.Drawing.Color]::White
$connectButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$connectButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
$connectButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$connectionCard.Controls.Add($connectButton)

# Création d'une carte pour les informations utilisateur
$userInfoCard = New-Object System.Windows.Forms.Panel
$userInfoCard.Location = New-Object System.Drawing.Point(30, 100)
$userInfoCard.Size = New-Object System.Drawing.Size(540, 420)
$userInfoCard.BackColor = $cardBgColor
$userInfoCard.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$mainContainer.Controls.Add($userInfoCard)

# Création d'une carte pour les options d'exportation
$exportCard = New-Object System.Windows.Forms.Panel
$exportCard.Location = New-Object System.Drawing.Point(30, 530)
$exportCard.Size = New-Object System.Drawing.Size(540, 50)
$exportCard.BackColor = $cardBgColor
$exportCard.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$mainContainer.Controls.Add($exportCard)

# Texte d'instruction pour l'exportation
$exportLabel = New-Object System.Windows.Forms.Label
$exportLabel.Location = New-Object System.Drawing.Point(20, 10)
$exportLabel.Size = New-Object System.Drawing.Size(200, 30)
$exportLabel.Text = "Exporter les données :"
$exportLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$exportLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$exportCard.Controls.Add($exportLabel)

# Bouton d'exportation JSON
$exportJsonButton = New-Object System.Windows.Forms.Button
$exportJsonButton.Location = New-Object System.Drawing.Point(350, 8)
$exportJsonButton.Size = New-Object System.Drawing.Size(170, 35)
$exportJsonButton.Text = "Exporter en JSON"
$exportJsonButton.BackColor = $primaryColor
$exportJsonButton.ForeColor = [System.Drawing.Color]::White
$exportJsonButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$exportJsonButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
$exportJsonButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$exportCard.Controls.Add($exportJsonButton)

# Création du TabControl pour les onglets
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(15, 15)
$tabControl.Size = New-Object System.Drawing.Size(510, 390)
$tabControl.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$userInfoCard.Controls.Add($tabControl)

# Création de l'onglet "Profil"
$tabProfil = New-Object System.Windows.Forms.TabPage
$tabProfil.Text = "Profil"
$tabProfil.BackColor = $cardBgColor
$tabControl.Controls.Add($tabProfil)

# Création de l'onglet "Tenant"
$tabTenant = New-Object System.Windows.Forms.TabPage
$tabTenant.Text = "Tenant"
$tabTenant.BackColor = $cardBgColor
$tabControl.Controls.Add($tabTenant)

# Création de l'onglet "Groupes"
$tabGroupes = New-Object System.Windows.Forms.TabPage
$tabGroupes.Text = "Groupes"
$tabGroupes.BackColor = $cardBgColor
$tabControl.Controls.Add($tabGroupes)

# Création de l'onglet "Recherche Utilisateur"
$tabRechercheUser = New-Object System.Windows.Forms.TabPage
$tabRechercheUser.Text = "Recherche Utilisateur"
$tabRechercheUser.BackColor = $cardBgColor
$tabControl.Controls.Add($tabRechercheUser)

# Zone de texte pour l'onglet Profil
$userInfoTextBox = New-Object System.Windows.Forms.RichTextBox
$userInfoTextBox.Location = New-Object System.Drawing.Point(10, 10)
$userInfoTextBox.Size = New-Object System.Drawing.Size(484, 342)
$userInfoTextBox.ReadOnly = $true
$userInfoTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$userInfoTextBox.BackColor = $cardBgColor
$userInfoTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$userInfoTextBox.ForeColor = $textColor
$userInfoTextBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
$tabProfil.Controls.Add($userInfoTextBox)

# Zone de texte pour l'onglet Tenant
$tenantInfoTextBox = New-Object System.Windows.Forms.RichTextBox
$tenantInfoTextBox.Location = New-Object System.Drawing.Point(10, 10)
$tenantInfoTextBox.Size = New-Object System.Drawing.Size(484, 342)
$tenantInfoTextBox.ReadOnly = $true
$tenantInfoTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$tenantInfoTextBox.BackColor = $cardBgColor
$tenantInfoTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$tenantInfoTextBox.ForeColor = $textColor
$tenantInfoTextBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
$tabTenant.Controls.Add($tenantInfoTextBox)

# Création du TreeView pour afficher les groupes et leurs membres
$groupsTreeView = New-Object System.Windows.Forms.TreeView
$groupsTreeView.Location = New-Object System.Drawing.Point(10, 10)
$groupsTreeView.Size = New-Object System.Drawing.Size(484, 342)
$groupsTreeView.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$groupsTreeView.BackColor = $cardBgColor
$groupsTreeView.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$groupsTreeView.ForeColor = $textColor
$groupsTreeView.Scrollable = $true
$groupsTreeView.ShowLines = $true
$groupsTreeView.ShowPlusMinus = $true
$groupsTreeView.ShowRootLines = $true
$groupsTreeView.Indent = 20
$tabGroupes.Controls.Add($groupsTreeView)

# Fonction pour ajouter du texte coloré au RichTextBox
function Add-ColoredText {
    param (
        [System.Windows.Forms.RichTextBox]$RichTextBox,
        [string]$Text,
        [System.Drawing.Color]$Color = [System.Drawing.Color]::Black,
        [System.Drawing.FontStyle]$Style = [System.Drawing.FontStyle]::Regular
    )
    
    if ($null -ne $RichTextBox -and $null -ne $Text) {
        # Convertir explicitement le texte en UTF-8 si nécessaire
        $encodedText = $Text
        
        $RichTextBox.SelectionStart = $RichTextBox.TextLength
        $RichTextBox.SelectionLength = 0
        
        # Vérifier que $Color n'est pas null avant de l'utiliser
        if ($null -ne $Color -and $Color -ne [System.Drawing.Color]::Empty) {
            $RichTextBox.SelectionColor = $Color
        } else {
            $RichTextBox.SelectionColor = $RichTextBox.ForeColor # Utiliser la couleur de texte par défaut
        }
        
        # Vérifier que $Style n'est pas null avant de créer la police
        if ($null -ne $Style) {
            $RichTextBox.SelectionFont = New-Object System.Drawing.Font("Segoe UI", 10, $Style)
        } else {
            $RichTextBox.SelectionFont = New-Object System.Drawing.Font("Segoe UI", 10)
        }
        
        $RichTextBox.AppendText($encodedText)
        
        # Faire défiler vers le haut pour afficher le début du texte
        $RichTextBox.SelectionStart = 0
        $RichTextBox.ScrollToCaret()
    }
}

# Fonction pour se connecter à Azure AD
$connectButton.Add_Click({
    $userInfoTextBox.Clear()
    $tenantInfoTextBox.Clear()
    
    # Si le TreeView existe, le vider aussi
    if ($null -ne $groupsTreeView) {
        $groupsTreeView.Nodes.Clear()
    }
    
    # Afficher un message de chargement avec animation
    Add-ColoredText -RichTextBox $userInfoTextBox -Text "Tentative de connexion a Azure AD..." -Color $primaryColor -Style Italic
    Add-ColoredText -RichTextBox $userInfoTextBox -Text ([Environment]::NewLine)
    
    # Changer le curseur pour indiquer le chargement
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    
    try {
        # Vérifier si le module AzureAD est disponible
        if (-not (Get-Module -ListAvailable -Name AzureAD)) {
            throw "Le module AzureAD n'est pas installé. Veuillez l'installer avec la commande: Install-Module -Name AzureAD -Force -AllowClobber"
        }
        
        # Vérifier si le module est chargé
        if (-not (Get-Module -Name AzureAD)) {
            Import-Module AzureAD -ErrorAction Stop
        }
        
        # Tentative de connexion à Azure AD - version simplifiée
        Connect-AzureAD -ErrorAction Stop
        
        # Si la connexion réussit (pas d'exception), récupérer les informations
        $currentUser = Get-AzureADCurrentSessionInfo -ErrorAction Stop
        $userDetails = Get-AzureADUser -ObjectId $currentUser.Account -ErrorAction Stop
        
        # Récupérer les informations du tenant
        $tenantDetails = Get-AzureADTenantDetail
        
        # Stocker les informations du tenant dans une variable globale pour l'exportation
        $script:tenantData = @{
            DisplayName = $tenantDetails.DisplayName
            TenantId = $currentUser.TenantId
            CountryLetterCode = $tenantDetails.CountryLetterCode
            PreferredLanguage = $tenantDetails.PreferredLanguage
            VerifiedDomains = ($tenantDetails.VerifiedDomains | ForEach-Object { 
                @{
                    Name = $_.Name
                    Default = $_.Default
                    Initial = $_.Initial
                    Type = $_.Type
                }
            })
            TechnicalContactEmail = $tenantDetails.TechnicalNotificationMails
        }
        
        # Stocker les informations utilisateur dans une variable globale pour l'exportation
        $script:userData = @{
            UserPrincipalName = $currentUser.Account
            DisplayName = $userDetails.DisplayName
            ObjectId = $userDetails.ObjectId
            Email = $userDetails.Mail
            OtherMails = $userDetails.OtherMails
            Department = $userDetails.Department
            JobTitle = $userDetails.JobTitle
            TelephoneNumber = $userDetails.TelephoneNumber
            Mobile = $userDetails.Mobile
            Office = $userDetails.PhysicalDeliveryOfficeName
            City = $userDetails.City
            Country = $userDetails.Country
            AccountEnabled = $userDetails.AccountEnabled
            CreatedDateTime = $userDetails.ExtensionProperty.createdDateTime
            LastSignIn = $userDetails.RefreshTokensValidFromDateTime
        }
        
        # Afficher les informations du tenant dans l'onglet Tenant
        $tenantInfoTextBox.Clear()
        Add-ColoredText -RichTextBox $tenantInfoTextBox -Text "Nom du Tenant : " -Color $primaryColor -Style Bold
        Add-ColoredText -RichTextBox $tenantInfoTextBox -Text ([string]$tenantDetails.DisplayName + [Environment]::NewLine + [Environment]::NewLine)
        
        Add-ColoredText -RichTextBox $tenantInfoTextBox -Text "ID du Tenant : " -Color $primaryColor -Style Bold
        Add-ColoredText -RichTextBox $tenantInfoTextBox -Text ([string]$currentUser.TenantId + [Environment]::NewLine + [Environment]::NewLine)
        
        Add-ColoredText -RichTextBox $tenantInfoTextBox -Text "Pays : " -Color $primaryColor -Style Bold
        Add-ColoredText -RichTextBox $tenantInfoTextBox -Text ([string]$tenantDetails.CountryLetterCode + [Environment]::NewLine + [Environment]::NewLine)
        
        Add-ColoredText -RichTextBox $tenantInfoTextBox -Text "Région : " -Color $primaryColor -Style Bold
        Add-ColoredText -RichTextBox $tenantInfoTextBox -Text ([string]$tenantDetails.PreferredLanguage + [Environment]::NewLine + [Environment]::NewLine)
        
        Add-ColoredText -RichTextBox $tenantInfoTextBox -Text "Domaines verifies : " -Color $primaryColor -Style Bold
        Add-ColoredText -RichTextBox $tenantInfoTextBox -Text ([Environment]::NewLine)
        
        foreach ($domain in $tenantDetails.VerifiedDomains) {
            Add-ColoredText -RichTextBox $tenantInfoTextBox -Text ("- " + $domain.Name)
            if ($domain.Default) {
                Add-ColoredText -RichTextBox $tenantInfoTextBox -Text " (Domaine par defaut)" -Color $accentColor
            }
            if ($domain.Initial) {
                Add-ColoredText -RichTextBox $tenantInfoTextBox -Text " (Domaine initial)" -Color $accentColor
            }
            Add-ColoredText -RichTextBox $tenantInfoTextBox -Text ([Environment]::NewLine)
        }
        Add-ColoredText -RichTextBox $tenantInfoTextBox -Text ([Environment]::NewLine)
        
        Add-ColoredText -RichTextBox $tenantInfoTextBox -Text "Contacts techniques : " -Color $primaryColor -Style Bold
        Add-ColoredText -RichTextBox $tenantInfoTextBox -Text ([Environment]::NewLine)
        
        if ($tenantDetails.TechnicalNotificationMails -and $tenantDetails.TechnicalNotificationMails.Count -gt 0) {
            foreach ($email in $tenantDetails.TechnicalNotificationMails) {
                Add-ColoredText -RichTextBox $tenantInfoTextBox -Text ("- " + $email + [Environment]::NewLine)
            }
        } else {
            Add-ColoredText -RichTextBox $tenantInfoTextBox -Text ("Aucun contact technique defini" + [Environment]::NewLine)
        }
        
        Add-ColoredText -RichTextBox $tenantInfoTextBox -Text ([Environment]::NewLine)
        Add-ColoredText -RichTextBox $tenantInfoTextBox -Text "Informations de securite : " -Color $primaryColor -Style Bold
        Add-ColoredText -RichTextBox $tenantInfoTextBox -Text ([Environment]::NewLine)
        
        # Récupérer les paramètres de mot de passe du tenant
        try {
            $passwordPolicy = Get-AzureADDirectorySettingTemplate | Where-Object {$_.DisplayName -eq "Password Rule Settings"}
            if ($passwordPolicy) {
                Add-ColoredText -RichTextBox $tenantInfoTextBox -Text ("- Politique de mot de passe configurée" + [Environment]::NewLine)
            } else {
                Add-ColoredText -RichTextBox $tenantInfoTextBox -Text ("- Politique de mot de passe par défaut" + [Environment]::NewLine)
            }
        } catch {
            Add-ColoredText -RichTextBox $tenantInfoTextBox -Text ("- Impossible de récupérer la politique de mot de passe" + [Environment]::NewLine)
        }
        
        # Récupérer les informations sur l'authentification multifacteur
        try {
            $mfaSettings = Get-AzureADMSAuthenticationMethodPolicyAssignment -ErrorAction SilentlyContinue
            if ($mfaSettings) {
                Add-ColoredText -RichTextBox $tenantInfoTextBox -Text ("- MFA configurée" + [Environment]::NewLine)
            } else {
                Add-ColoredText -RichTextBox $tenantInfoTextBox -Text ("- MFA non configurée ou impossible à déterminer" + [Environment]::NewLine)
            }
        } catch {
            Add-ColoredText -RichTextBox $tenantInfoTextBox -Text ("- Impossible de récupérer les paramètres MFA" + [Environment]::NewLine)
        }
        
        # Stocker les informations utilisateur dans une variable globale pour l'exportation
        $script:userData = @{
            UserPrincipalName = $currentUser.Account
            DisplayName = $userDetails.DisplayName
            ObjectId = $userDetails.ObjectId
            Email = $userDetails.Mail
            Department = $userDetails.Department
            JobTitle = $userDetails.JobTitle
            TelephoneNumber = $userDetails.TelephoneNumber
        }
        
        # Activer les boutons d'exportation
        $exportJsonButton.Enabled = $true
        
        # Effacer le contenu précédent
        # Effacer le contenu précédent
        $userInfoTextBox.Clear()
        
        # Informations utilisateur avec mise en forme
        Add-ColoredText -RichTextBox $userInfoTextBox -Text "Nom d'utilisateur : " -Color $primaryColor -Style Bold
        Add-ColoredText -RichTextBox $userInfoTextBox -Text ([string]$currentUser.Account + [Environment]::NewLine)
        
        Add-ColoredText -RichTextBox $userInfoTextBox -Text "Nom complet : " -Color $primaryColor -Style Bold
        Add-ColoredText -RichTextBox $userInfoTextBox -Text ([string]$userDetails.DisplayName + [Environment]::NewLine)
        
        Add-ColoredText -RichTextBox $userInfoTextBox -Text "ID de l'objet : " -Color $primaryColor -Style Bold
        Add-ColoredText -RichTextBox $userInfoTextBox -Text ([string]$userDetails.ObjectId + [Environment]::NewLine)
        
        Add-ColoredText -RichTextBox $userInfoTextBox -Text "Email : " -Color $primaryColor -Style Bold
        Add-ColoredText -RichTextBox $userInfoTextBox -Text ([string]$userDetails.Mail + [Environment]::NewLine)
        
        Add-ColoredText -RichTextBox $userInfoTextBox -Text "Email secondaire : " -Color $primaryColor -Style Bold
        Add-ColoredText -RichTextBox $userInfoTextBox -Text ([string]$userDetails.OtherMails -join ", " + [Environment]::NewLine)
        
        Add-ColoredText -RichTextBox $userInfoTextBox -Text "Departement : " -Color $primaryColor -Style Bold
        Add-ColoredText -RichTextBox $userInfoTextBox -Text ([string]$userDetails.Department + [Environment]::NewLine)
        
        Add-ColoredText -RichTextBox $userInfoTextBox -Text "Titre : " -Color $primaryColor -Style Bold
        Add-ColoredText -RichTextBox $userInfoTextBox -Text ([string]$userDetails.JobTitle + [Environment]::NewLine)
        
        Add-ColoredText -RichTextBox $userInfoTextBox -Text "Telephone : " -Color $primaryColor -Style Bold
        Add-ColoredText -RichTextBox $userInfoTextBox -Text ([string]$userDetails.TelephoneNumber + [Environment]::NewLine)
        
        Add-ColoredText -RichTextBox $userInfoTextBox -Text "Mobile : " -Color $primaryColor -Style Bold
        Add-ColoredText -RichTextBox $userInfoTextBox -Text ([string]$userDetails.Mobile + [Environment]::NewLine)
        
        Add-ColoredText -RichTextBox $userInfoTextBox -Text "Bureau : " -Color $primaryColor -Style Bold
        Add-ColoredText -RichTextBox $userInfoTextBox -Text ([string]$userDetails.PhysicalDeliveryOfficeName + [Environment]::NewLine)
        
        Add-ColoredText -RichTextBox $userInfoTextBox -Text "Ville : " -Color $primaryColor -Style Bold
        Add-ColoredText -RichTextBox $userInfoTextBox -Text ([string]$userDetails.City + [Environment]::NewLine)
        
        Add-ColoredText -RichTextBox $userInfoTextBox -Text "Pays : " -Color $primaryColor -Style Bold
        Add-ColoredText -RichTextBox $userInfoTextBox -Text ([string]$userDetails.Country + [Environment]::NewLine)
        
        Add-ColoredText -RichTextBox $userInfoTextBox -Text "Statut du compte : " -Color $primaryColor -Style Bold
        if ($userDetails.AccountEnabled) {
            Add-ColoredText -RichTextBox $userInfoTextBox -Text ("Actif" + [Environment]::NewLine) -Color $accentColor
        } else {
            Add-ColoredText -RichTextBox $userInfoTextBox -Text ("Inactif" + [Environment]::NewLine) -Color $errorColor
        }
        
        Add-ColoredText -RichTextBox $userInfoTextBox -Text "Date de création : " -Color $primaryColor -Style Bold
        Add-ColoredText -RichTextBox $userInfoTextBox -Text ([string]$userDetails.ExtensionProperty.createdDateTime + [Environment]::NewLine)
        
        Add-ColoredText -RichTextBox $userInfoTextBox -Text "Dernière connexion : " -Color $primaryColor -Style Bold
        Add-ColoredText -RichTextBox $userInfoTextBox -Text ([string]$userDetails.RefreshTokensValidFromDateTime + [Environment]::NewLine + [Environment]::NewLine)
        
        # Récupérer les groupes de l'utilisateur
        $script:userGroups = Get-AzureADUserMembership -ObjectId $userDetails.ObjectId | Where-Object { $_.ObjectType -eq "Group" }
        
        # Afficher les groupes dans l'onglet Profil
        Add-ColoredText -RichTextBox $userInfoTextBox -Text "Groupes : " -Color $primaryColor -Style Bold
        Add-ColoredText -RichTextBox $userInfoTextBox -Text ([Environment]::NewLine)
        
        foreach ($group in $script:userGroups) {
            Add-ColoredText -RichTextBox $userInfoTextBox -Text ([string]$group.DisplayName + [Environment]::NewLine)
            
            # Si le TreeView existe, ajouter les groupes
            if ($null -ne $groupsTreeView) {
                $groupNode = New-Object System.Windows.Forms.TreeNode
                $groupNode.Text = $group.DisplayName
                $groupNode.ForeColor = $primaryColor
                $groupNode.NodeFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
                $groupNode.Tag = $group.ObjectId  # Stocker l'ID du groupe dans la propriété Tag
                $groupsTreeView.Nodes.Add($groupNode)
                
                # Ajouter un nœud de chargement temporaire
                $loadingNode = New-Object System.Windows.Forms.TreeNode
                $loadingNode.Text = "Chargement des membres..."
                $loadingNode.ForeColor = $textColor
                $groupNode.Nodes.Add($loadingNode)
            }
        }
        
        # Configurer l'événement BeforeExpand pour le TreeView à l'intérieur du gestionnaire de clic
        # Utiliser la méthode add_BeforeExpand au lieu de la propriété BeforeExpand
        $groupsTreeView.add_BeforeExpand({
            param($sender, $e)
            
            $selectedNode = $e.Node
            
            # Vérifier si c'est la première expansion (nœud de chargement présent)
            if ($selectedNode.Nodes.Count -eq 1 -and $selectedNode.Nodes[0].Text -eq "Chargement des membres...") {
                # Supprimer le nœud de chargement
                $selectedNode.Nodes.Clear()
                
                # Récupérer l'ID du groupe à partir de la propriété Tag
                $groupId = $selectedNode.Tag
                
                if ($groupId) {
                    # Récupérer les membres du groupe
                    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
                    try {
                        $groupMembers = Get-AzureADGroupMember -ObjectId $groupId -All $true
                        
                        # Ajouter chaque membre au nœud du groupe
                        foreach ($member in $groupMembers) {
                            $memberNode = New-Object System.Windows.Forms.TreeNode
                            
                            # Déterminer le type de membre et formater en conséquence
                            if ($member.ObjectType -eq "User") {
                                $memberNode.Text = "$($member.DisplayName) ($($member.UserPrincipalName))"
                                $memberNode.ForeColor = $textColor
                            } else {
                                $memberNode.Text = "$($member.DisplayName) (Groupe)"
                                $memberNode.ForeColor = $accentColor
                            }
                            
                            $selectedNode.Nodes.Add($memberNode)
                        }
                        
                        # Si aucun membre n'est trouvé
                        if ($groupMembers.Count -eq 0) {
                            $emptyNode = New-Object System.Windows.Forms.TreeNode
                            $emptyNode.Text = "Aucun membre"
                            $emptyNode.ForeColor = $errorColor
                            $selectedNode.Nodes.Add($emptyNode)
                        }
                    }
                    catch {
                        # En cas d'erreur lors de la récupération des membres
                        $errorNode = New-Object System.Windows.Forms.TreeNode
                        $errorNode.Text = "Erreur: $($_.Exception.Message)"
                        $errorNode.ForeColor = $errorColor
                        $selectedNode.Nodes.Add($errorNode)
                    }
                    finally {
                        $form.Cursor = [System.Windows.Forms.Cursors]::Default
                    }
                }
            }
        })
    }
    catch {
        # En cas d'erreur, afficher un message d'erreur détaillé
        $userInfoTextBox.Clear()
        Add-ColoredText -RichTextBox $userInfoTextBox -Text "❌ Erreur de connexion : " -Color $errorColor -Style Bold
        Add-ColoredText -RichTextBox $userInfoTextBox -Text ($_.Exception.Message + [Environment]::NewLine)
        
        # Désactiver les boutons d'exportation
        & $script:DisableExportButtons
        
        # Ajouter des informations de débogage supplémentaires
        Add-ColoredText -RichTextBox $userInfoTextBox -Text ([Environment]::NewLine + "Détails techniques pour le dépannage :" + [Environment]::NewLine) -Color $primaryColor -Style Bold
        
        # Vérifier si le module est installé
        if (-not (Get-Module -ListAvailable -Name AzureAD)) {
            Add-ColoredText -RichTextBox $userInfoTextBox -Text "- Le module AzureAD n'est pas installé." -Color $errorColor
            Add-ColoredText -RichTextBox $userInfoTextBox -Text ([Environment]::NewLine)
            Add-ColoredText -RichTextBox $userInfoTextBox -Text "  Exécutez cette commande pour l'installer :" -Color $textColor
            Add-ColoredText -RichTextBox $userInfoTextBox -Text ([Environment]::NewLine)
            Add-ColoredText -RichTextBox $userInfoTextBox -Text "  Install-Module -Name AzureAD -Force -AllowClobber" -Color $primaryColor
            Add-ColoredText -RichTextBox $userInfoTextBox -Text ([Environment]::NewLine + [Environment]::NewLine)
        } else {
            Add-ColoredText -RichTextBox $userInfoTextBox -Text "- Le module AzureAD est correctement installé." -Color $accentColor
            Add-ColoredText -RichTextBox $userInfoTextBox -Text ([Environment]::NewLine + [Environment]::NewLine)
        }
        
        # Vérifier la connectivité Internet
        try {
            $internetTest = Test-Connection -ComputerName "login.microsoftonline.com" -Count 1 -Quiet
            if ($internetTest) {
                Add-ColoredText -RichTextBox $userInfoTextBox -Text "- Connexion Internet : OK" -Color $accentColor
            } else {
                Add-ColoredText -RichTextBox $userInfoTextBox -Text "- Connexion Internet : Problème détecté" -Color $errorColor
            }
            Add-ColoredText -RichTextBox $userInfoTextBox -Text ([Environment]::NewLine + [Environment]::NewLine)
        } catch {
            Add-ColoredText -RichTextBox $userInfoTextBox -Text "- Connexion Internet : Impossible de vérifier" -Color $errorColor
            Add-ColoredText -RichTextBox $userInfoTextBox -Text ([Environment]::NewLine + [Environment]::NewLine)
        }
        
        # Conseils de dépannage
        Add-ColoredText -RichTextBox $userInfoTextBox -Text "Suggestions :" -Color $primaryColor -Style Bold
        Add-ColoredText -RichTextBox $userInfoTextBox -Text ([Environment]::NewLine)
        Add-ColoredText -RichTextBox $userInfoTextBox -Text "1. Vérifiez vos identifiants Azure AD" -Color $textColor
        Add-ColoredText -RichTextBox $userInfoTextBox -Text ([Environment]::NewLine)
        Add-ColoredText -RichTextBox $userInfoTextBox -Text "2. Assurez-vous d'avoir une connexion Internet active" -Color $textColor
        Add-ColoredText -RichTextBox $userInfoTextBox -Text ([Environment]::NewLine)
        Add-ColoredText -RichTextBox $userInfoTextBox -Text "3. Vérifiez que vous avez les autorisations nécessaires" -Color $textColor
        Add-ColoredText -RichTextBox $userInfoTextBox -Text ([Environment]::NewLine)
        Add-ColoredText -RichTextBox $userInfoTextBox -Text "4. Redémarrez PowerShell et réessayez" -Color $textColor
    }
    finally {
        # Restaurer le curseur
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# Interface de recherche d'utilisateur
$searchUserLabel = New-Object System.Windows.Forms.Label
$searchUserLabel.Location = New-Object System.Drawing.Point(10, 15)
$searchUserLabel.Size = New-Object System.Drawing.Size(220, 25)
$searchUserLabel.Text = "Rechercher un utilisateur:"
$searchUserLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$tabRechercheUser.Controls.Add($searchUserLabel)

$searchUserTextBox = New-Object System.Windows.Forms.TextBox
$searchUserTextBox.Location = New-Object System.Drawing.Point(10, 45)
$searchUserTextBox.Size = New-Object System.Drawing.Size(350, 25)
$searchUserTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$tabRechercheUser.Controls.Add($searchUserTextBox)

$searchUserButton = New-Object System.Windows.Forms.Button
$searchUserButton.Location = New-Object System.Drawing.Point(370, 45)
$searchUserButton.Size = New-Object System.Drawing.Size(100, 25)
$searchUserButton.Text = "Rechercher"
$searchUserButton.BackColor = $primaryColor
$searchUserButton.ForeColor = [System.Drawing.Color]::White
$searchUserButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$tabRechercheUser.Controls.Add($searchUserButton)

# Creation d'une liste pour afficher les resultats de recherche
$userResultsListView = New-Object System.Windows.Forms.ListView
$userResultsListView.Location = New-Object System.Drawing.Point(10, 80)
$userResultsListView.Size = New-Object System.Drawing.Size(484, 200)
$userResultsListView.View = [System.Windows.Forms.View]::Details
$userResultsListView.FullRowSelect = $true
$userResultsListView.GridLines = $true
$userResultsListView.BackColor = $cardBgColor
$userResultsListView.ForeColor = $textColor
$userResultsListView.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$tabRechercheUser.Controls.Add($userResultsListView)

# Ajouter les colonnes
$userResultsListView.Columns.Add("Nom", 150)
$userResultsListView.Columns.Add("Email", 150)
$userResultsListView.Columns.Add("Titre", 180)

# Zone de details pour l'utilisateur selectionne
$userDetailsBox = New-Object System.Windows.Forms.RichTextBox
$userDetailsBox.Location = New-Object System.Drawing.Point(10, 290)
$userDetailsBox.Size = New-Object System.Drawing.Size(484, 60)
$userDetailsBox.ReadOnly = $true
$userDetailsBox.BackColor = $cardBgColor
$userDetailsBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$userDetailsBox.ForeColor = $textColor
$userDetailsBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$tabRechercheUser.Controls.Add($userDetailsBox)

# Permettre la recherche en appuyant sur Entrée dans le champ de recherche
$searchUserTextBox.Add_KeyDown({
    param($sender, $e)
    
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $searchUserButton.PerformClick()
        $e.SuppressKeyPress = $true
    }
})

# Gestionnaire d'evenements pour le bouton de recherche
$searchUserButton.Add_Click({
    $searchTerm = $searchUserTextBox.Text.Trim()
    
    if ($searchTerm -ne "") {
        # Effacer les resultats precedents
        $userResultsListView.Items.Clear()
        $userDetailsBox.Clear()
        
        # Changer le curseur pour indiquer le chargement
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        
        try {
            # Verifier si nous sommes connectes a Azure AD
            try {
                $currentSession = Get-AzureADCurrentSessionInfo -ErrorAction Stop
            }
            catch {
                throw "Vous n'etes pas connecte a Azure AD. Veuillez vous connecter d'abord."
            }
            
            # Rechercher les utilisateurs avec une methode plus flexible
            try {
                # Recuperer tous les utilisateurs et filtrer cote client
                # Cette methode est plus lente mais permet une recherche plus flexible
                $users = Get-AzureADUser -All $true -ErrorAction Stop | 
                         Where-Object { 
                             $_.DisplayName -like "*$searchTerm*" -or 
                             $_.GivenName -like "*$searchTerm*" -or 
                             $_.Surname -like "*$searchTerm*" -or 
                             $_.UserPrincipalName -like "*$searchTerm*" -or 
                             $_.Mail -like "*$searchTerm*" -or
                             $_.JobTitle -like "*$searchTerm*" 
                         } | 
                         Select-Object -First 100
            }
            catch {
                # Si la methode precedente echoue, essayer avec -Filter
                try {
                    $filter = "substringof('$searchTerm', DisplayName) or substringof('$searchTerm', UserPrincipalName) or substringof('$searchTerm', Mail)"
                    $users = Get-AzureADUser -Filter $filter -Top 100 -ErrorAction Stop
                }
                catch {
                    # Si tout echoue, utiliser la methode la plus simple
                    $users = Get-AzureADUser -SearchString $searchTerm -Top 100 -ErrorAction Stop
                }
            }
            
            if ($users.Count -eq 0) {
                $userDetailsBox.Clear()
                Add-ColoredText -RichTextBox $userDetailsBox -Text "Aucun utilisateur trouve pour la recherche: " -Color $errorColor -Style Bold
                Add-ColoredText -RichTextBox $userDetailsBox -Text $searchTerm
            }
            else {
                # Ajouter les utilisateurs a la liste
                foreach ($user in $users) {
                    $item = New-Object System.Windows.Forms.ListViewItem($user.DisplayName)
                    $item.SubItems.Add($user.Mail)
                    $item.SubItems.Add($user.JobTitle)
                    $item.Tag = $user.ObjectId  # Stocker l'ID pour une utilisation ulterieure
                    $userResultsListView.Items.Add($item)
                }
                
                # Afficher le nombre de resultats
                $userDetailsBox.Clear()
                Add-ColoredText -RichTextBox $userDetailsBox -Text "$($users.Count) utilisateur(s) trouve(s) pour la recherche: " -Color $accentColor -Style Bold
                Add-ColoredText -RichTextBox $userDetailsBox -Text $searchTerm
            }
        }
        catch {
            $userDetailsBox.Clear()
            Add-ColoredText -RichTextBox $userDetailsBox -Text "Erreur: " -Color $errorColor -Style Bold
            Add-ColoredText -RichTextBox $userDetailsBox -Text $_.Exception.Message
        }
        finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }
    else {
        $userDetailsBox.Clear()
        Add-ColoredText -RichTextBox $userDetailsBox -Text "Veuillez entrer un terme de recherche." -Color $errorColor
    }
})

# Gestionnaire d'evenements pour la selection d'un utilisateur dans la liste
$userResultsListView.Add_SelectedIndexChanged({
    if ($userResultsListView.SelectedItems.Count -gt 0) {
        $selectedUserId = $userResultsListView.SelectedItems[0].Tag
        
        if ($selectedUserId) {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            
            try {
                # Recuperer les details complets de l'utilisateur
                $userDetails = Get-AzureADUser -ObjectId $selectedUserId
                
                # Creer une fenetre de details
                $detailsForm = New-Object System.Windows.Forms.Form
                $detailsForm.Text = "Details de l'utilisateur: $($userDetails.DisplayName)"
                $detailsForm.Size = New-Object System.Drawing.Size(500, 600)
                $detailsForm.StartPosition = "CenterParent"
                $detailsForm.BackColor = $bgColor
                $detailsForm.Font = New-Object System.Drawing.Font("Segoe UI", 10)
                
                # Creation d'un RichTextBox pour afficher les details
                $detailsTextBox = New-Object System.Windows.Forms.RichTextBox
                $detailsTextBox.Location = New-Object System.Drawing.Point(10, 10)
                $detailsTextBox.Size = New-Object System.Drawing.Size(465, 300)
                $detailsTextBox.ReadOnly = $true
                $detailsTextBox.BackColor = $cardBgColor
                $detailsTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
                $detailsTextBox.ForeColor = $textColor
                $detailsForm.Controls.Add($detailsTextBox)
                
                # Afficher les informations detaillees
                Add-ColoredText -RichTextBox $detailsTextBox -Text "Nom d'utilisateur: " -Color $primaryColor -Style Bold
                Add-ColoredText -RichTextBox $detailsTextBox -Text ([string]$userDetails.UserPrincipalName + [Environment]::NewLine)
                
                Add-ColoredText -RichTextBox $detailsTextBox -Text "Nom complet: " -Color $primaryColor -Style Bold
                Add-ColoredText -RichTextBox $detailsTextBox -Text ([string]$userDetails.DisplayName + [Environment]::NewLine)
                
                Add-ColoredText -RichTextBox $detailsTextBox -Text "ID de l'objet: " -Color $primaryColor -Style Bold
                Add-ColoredText -RichTextBox $detailsTextBox -Text ([string]$userDetails.ObjectId + [Environment]::NewLine)
                
                Add-ColoredText -RichTextBox $detailsTextBox -Text "Email: " -Color $primaryColor -Style Bold
                Add-ColoredText -RichTextBox $detailsTextBox -Text ([string]$userDetails.Mail + [Environment]::NewLine)
                
                Add-ColoredText -RichTextBox $detailsTextBox -Text "Email secondaire: " -Color $primaryColor -Style Bold
                Add-ColoredText -RichTextBox $detailsTextBox -Text ([string]$userDetails.OtherMails -join ", " + [Environment]::NewLine)
                
                Add-ColoredText -RichTextBox $detailsTextBox -Text "Departement: " -Color $primaryColor -Style Bold
                Add-ColoredText -RichTextBox $detailsTextBox -Text ([string]$userDetails.Department + [Environment]::NewLine)
                
                Add-ColoredText -RichTextBox $detailsTextBox -Text "Titre: " -Color $primaryColor -Style Bold
                Add-ColoredText -RichTextBox $detailsTextBox -Text ([string]$userDetails.JobTitle + [Environment]::NewLine)
                
                Add-ColoredText -RichTextBox $detailsTextBox -Text "Telephone: " -Color $primaryColor -Style Bold
                Add-ColoredText -RichTextBox $detailsTextBox -Text ([string]$userDetails.TelephoneNumber + [Environment]::NewLine + [Environment]::NewLine)
                
                Add-ColoredText -RichTextBox $detailsTextBox -Text "Mobile: " -Color $primaryColor -Style Bold
                Add-ColoredText -RichTextBox $detailsTextBox -Text ([string]$userDetails.Mobile + [Environment]::NewLine)
                
                Add-ColoredText -RichTextBox $detailsTextBox -Text "Bureau: " -Color $primaryColor -Style Bold
                Add-ColoredText -RichTextBox $detailsTextBox -Text ([string]$userDetails.PhysicalDeliveryOfficeName + [Environment]::NewLine)
                
                Add-ColoredText -RichTextBox $detailsTextBox -Text "Ville: " -Color $primaryColor -Style Bold
                Add-ColoredText -RichTextBox $detailsTextBox -Text ([string]$userDetails.City + [Environment]::NewLine)
                
                Add-ColoredText -RichTextBox $detailsTextBox -Text "Pays: " -Color $primaryColor -Style Bold
                Add-ColoredText -RichTextBox $detailsTextBox -Text ([string]$userDetails.Country + [Environment]::NewLine)
                
                Add-ColoredText -RichTextBox $detailsTextBox -Text "Statut du compte: " -Color $primaryColor -Style Bold
                if ($userDetails.AccountEnabled) {
                    Add-ColoredText -RichTextBox $detailsTextBox -Text ("Actif" + [Environment]::NewLine) -Color $accentColor
                } else {
                    Add-ColoredText -RichTextBox $detailsTextBox -Text ("Inactif" + [Environment]::NewLine) -Color $errorColor
                }
                
                Add-ColoredText -RichTextBox $detailsTextBox -Text "Date de création: " -Color $primaryColor -Style Bold
                Add-ColoredText -RichTextBox $detailsTextBox -Text ([string]$userDetails.ExtensionProperty.createdDateTime + [Environment]::NewLine)
                
                Add-ColoredText -RichTextBox $detailsTextBox -Text "Dernière connexion: " -Color $primaryColor -Style Bold
                Add-ColoredText -RichTextBox $detailsTextBox -Text ([string]$userDetails.RefreshTokensValidFromDateTime + [Environment]::NewLine + [Environment]::NewLine)
                
                # Récupérer les groupes de l'utilisateur
                $script:userGroups = Get-AzureADUserMembership -ObjectId $userDetails.ObjectId | Where-Object { $_.ObjectType -eq "Group" }
                
                # Afficher les groupes dans l'onglet Profil
                Add-ColoredText -RichTextBox $detailsTextBox -Text "Groupes: " -Color $primaryColor -Style Bold
                Add-ColoredText -RichTextBox $detailsTextBox -Text ([Environment]::NewLine)
                
                try {
                    $userGroups = Get-AzureADUserMembership -ObjectId $userDetails.ObjectId | Where-Object { $_.ObjectType -eq "Group" }
                    
                    if ($userGroups.Count -eq 0) {
                        Add-ColoredText -RichTextBox $detailsTextBox -Text "Aucun groupe trouve" -Color $textColor
                    }
                    else {
                        foreach ($group in $userGroups) {
                            Add-ColoredText -RichTextBox $detailsTextBox -Text "- " -Color $accentColor -Style Bold
                            Add-ColoredText -RichTextBox $detailsTextBox -Text ([string]$group.DisplayName + [Environment]::NewLine)
                        }
                    }
                }
                catch {
                    Add-ColoredText -RichTextBox $detailsTextBox -Text "Erreur lors de la recuperation des groupes: $($_.Exception.Message)" -Color $errorColor
                }
                
                # Creation d'un ListView pour les licences
                $licensesLabel = New-Object System.Windows.Forms.Label
                $licensesLabel.Location = New-Object System.Drawing.Point(10, 320)
                $licensesLabel.Size = New-Object System.Drawing.Size(465, 25)
                $licensesLabel.Text = "Licences attribuees:"
                $licensesLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
                $licensesLabel.ForeColor = $primaryColor
                $detailsForm.Controls.Add($licensesLabel)
                
                $licensesListView = New-Object System.Windows.Forms.ListView
                $licensesListView.Location = New-Object System.Drawing.Point(10, 350)
                $licensesListView.Size = New-Object System.Drawing.Size(465, 150)
                $licensesListView.View = [System.Windows.Forms.View]::Details
                $licensesListView.FullRowSelect = $true
                $licensesListView.GridLines = $true
                $licensesListView.BackColor = $cardBgColor
                $licensesListView.ForeColor = $textColor
                $licensesListView.Font = New-Object System.Drawing.Font("Segoe UI", 9)
                $detailsForm.Controls.Add($licensesListView)
                
                # Ajouter les colonnes pour les licences
                $licensesListView.Columns.Add("Nom de la licence", 300)
                $licensesListView.Columns.Add("Etat", 150)
                
                # Recuperer et afficher les licences
                try {
                    $assignedLicenses = Get-AzureADUserLicenseDetail -ObjectId $userDetails.ObjectId
                    
                    if ($assignedLicenses.Count -eq 0) {
                        $item = New-Object System.Windows.Forms.ListViewItem("Aucune licence attribuee")
                        $licensesListView.Items.Add($item)
                    }
                    else {
                        foreach ($license in $assignedLicenses) {
                            $item = New-Object System.Windows.Forms.ListViewItem($license.SkuPartNumber)
                            $item.SubItems.Add("Attribuee")
                            $licensesListView.Items.Add($item)
                        }
                    }
                }
                catch {
                    $item = New-Object System.Windows.Forms.ListViewItem("Erreur: $($_.Exception.Message)")
                    $licensesListView.Items.Add($item)
                }
                
                # Bouton de fermeture
                $closeButton = New-Object System.Windows.Forms.Button
                $closeButton.Location = New-Object System.Drawing.Point(200, 520)
                $closeButton.Size = New-Object System.Drawing.Size(100, 30)
                $closeButton.Text = "Fermer"
                $closeButton.BackColor = $primaryColor
                $closeButton.ForeColor = [System.Drawing.Color]::White
                $closeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                $closeButton.Add_Click({ $detailsForm.Close() })
                $detailsForm.Controls.Add($closeButton)
                
                # Afficher la fenetre de details
                [void]$detailsForm.ShowDialog()
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show(
                    "Erreur lors de la recuperation des details de l'utilisateur: $($_.Exception.Message)",
                    "Erreur",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
            finally {
                $form.Cursor = [System.Windows.Forms.Cursors]::Default
            }
        }
    }
})

# Fonction pour exporter les données en JSON
$exportJsonButton.Add_Click({
    try {
        # Créer un objet avec toutes les données à exporter
        $exportData = @{
            TenantInfo = $script:tenantData
            UserInfo = $script:userData
            Groups = @()
        }
        
        # Ajouter les informations des groupes
        foreach ($group in $script:userGroups) {
            $groupInfo = @{
                DisplayName = $group.DisplayName
                Description = $group.Description
                ObjectId = $group.ObjectId
                ObjectType = $group.ObjectType
            }
            $exportData.Groups += $groupInfo
        }
        
        # Créer une boîte de dialogue pour enregistrer le fichier
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "Fichiers JSON (*.json)|*.json"
        $saveFileDialog.Title = "Enregistrer les données en JSON"
        $saveFileDialog.FileName = "$($script:userData.DisplayName)_AzureAD_Export.json"
        
        # Si l'utilisateur clique sur "Enregistrer"
        if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            # Convertir les données en JSON et les enregistrer
            $jsonData = $exportData | ConvertTo-Json -Depth 4
            $jsonData | Out-File -FilePath $saveFileDialog.FileName -Encoding UTF8
            
            # Afficher un message de confirmation
            [System.Windows.Forms.MessageBox]::Show(
                "Les données ont été exportées avec succès vers $($saveFileDialog.FileName)",
                "Exportation réussie",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
    }
    catch {
        # Afficher un message d'erreur en cas de problème
        [System.Windows.Forms.MessageBox]::Show(
            "Erreur lors de l'exportation des données : $($_.Exception.Message)",
            "Erreur d'exportation",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
})

# Ajuster la taille du formulaire pour accommoder la nouvelle carte d'exportation
$form.Size = New-Object System.Drawing.Size(600, 700)  # Augmentation de la hauteur de 600 à 700
$mainContainer.Size = New-Object System.Drawing.Size(600, 630)  # Augmentation de la hauteur pour accommoder la carte d'exportation

# Désactiver les boutons d'exportation en cas de déconnexion ou d'erreur
$script:DisableExportButtons = {
    $exportJsonButton.Enabled = $falsee
}

# Modifier le gestionnaire d'erreurs pour désactiver les boutons d'exportation
# Dans le bloc catch de la fonction de connexion, ajouter:
# & $script:DisableExportButtons

# Affichage de la fenetre
[void]$form.ShowDialog()




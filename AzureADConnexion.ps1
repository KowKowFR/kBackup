[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Importation des modules nécessaires
# Si le module AzureAD n'est pas installé, vous devrez l'installer avec :
# Install-Module -Name AzureAD -Force -AllowClobber

# Chargement des assemblies pour l'interface graphique
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Définition des couleurs pour une interface moderne
$primaryColor = [System.Drawing.Color]::FromArgb(0, 120, 212)     # Bleu Microsoft
$accentColor = [System.Drawing.Color]::FromArgb(0, 200, 83)       # Vert pour succès
$errorColor = [System.Drawing.Color]::FromArgb(209, 52, 56)       # Rouge pour erreur
$textColor = [System.Drawing.Color]::FromArgb(50, 49, 48)         # Gris foncé pour texte
$bgColor = [System.Drawing.Color]::FromArgb(250, 250, 250)        # Fond légèrement grisé
$cardBgColor = [System.Drawing.Color]::White                      # Fond blanc pour les cartes

# Création de la fenêtre principale avec un design plus moderne
$form = New-Object System.Windows.Forms.Form
$form.Text = "Azure AD - Interface Graphique"
$form.Size = New-Object System.Drawing.Size(600, 600)  # Augmentation de la hauteur de 500 à 600
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

# Création du titre dans l'en-tête
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location = New-Object System.Drawing.Point(20, 20)
$titleLabel.Size = New-Object System.Drawing.Size(520, 30)
$titleLabel.Text = "Azure Active Directory Interface"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::White
$titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$headerPanel.Controls.Add($titleLabel)



# Création d'un conteneur principal pour le contenu
$mainContainer = New-Object System.Windows.Forms.Panel
$mainContainer.Location = New-Object System.Drawing.Point(0, 70)
$mainContainer.Size = New-Object System.Drawing.Size(600, 530)  # Augmentation de la hauteur de 430 à 530
$mainContainer.BackColor = $bgColor
$mainContainer.Padding = New-Object System.Windows.Forms.Padding(20, 20, 20, 20)
$form.Controls.Add($mainContainer)

# Création d'une carte pour le bouton de connexion
$connectionCard = New-Object System.Windows.Forms.Panel
$connectionCard.Location = New-Object System.Drawing.Point(30, 30)
$connectionCard.Size = New-Object System.Drawing.Size(540, 80)
$connectionCard.BackColor = $cardBgColor
$connectionCard.BorderStyle = [System.Windows.Forms.BorderStyle]::None
# Ajouter une ombre (simulée avec une bordure)
$connectionCard.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$mainContainer.Controls.Add($connectionCard)

# Texte d'instruction
$instructionLabel = New-Object System.Windows.Forms.Label
$instructionLabel.Location = New-Object System.Drawing.Point(20, 15)
$instructionLabel.Size = New-Object System.Drawing.Size(300, 50)
$instructionLabel.Text = "Connectez-vous pour acceder a vos informations Azure AD"
$instructionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$instructionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$connectionCard.Controls.Add($instructionLabel)

# Création du bouton de connexion avec style moderne
$connectButton = New-Object System.Windows.Forms.Button
$connectButton.Location = New-Object System.Drawing.Point(350, 20)
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
$userInfoCard.Location = New-Object System.Drawing.Point(30, 130)
$userInfoCard.Size = New-Object System.Drawing.Size(540, 420)  # Augmentation de la hauteur de 320 à 420
$userInfoCard.BackColor = $cardBgColor
$userInfoCard.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$mainContainer.Controls.Add($userInfoCard)

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
        
        Add-ColoredText -RichTextBox $userInfoTextBox -Text "Departement : " -Color $primaryColor -Style Bold
        Add-ColoredText -RichTextBox $userInfoTextBox -Text ([string]$userDetails.Department + [Environment]::NewLine)
        
        Add-ColoredText -RichTextBox $userInfoTextBox -Text "Titre : " -Color $primaryColor -Style Bold
        Add-ColoredText -RichTextBox $userInfoTextBox -Text ([string]$userDetails.JobTitle + [Environment]::NewLine)
        
        Add-ColoredText -RichTextBox $userInfoTextBox -Text "Telephone : " -Color $primaryColor -Style Bold
        Add-ColoredText -RichTextBox $userInfoTextBox -Text ([string]$userDetails.TelephoneNumber + [Environment]::NewLine + [Environment]::NewLine)
        
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
                
                Add-ColoredText -RichTextBox $detailsTextBox -Text "Departement: " -Color $primaryColor -Style Bold
                Add-ColoredText -RichTextBox $detailsTextBox -Text ([string]$userDetails.Department + [Environment]::NewLine)
                
                Add-ColoredText -RichTextBox $detailsTextBox -Text "Titre: " -Color $primaryColor -Style Bold
                Add-ColoredText -RichTextBox $detailsTextBox -Text ([string]$userDetails.JobTitle + [Environment]::NewLine)
                
                Add-ColoredText -RichTextBox $detailsTextBox -Text "Telephone: " -Color $primaryColor -Style Bold
                Add-ColoredText -RichTextBox $detailsTextBox -Text ([string]$userDetails.TelephoneNumber + [Environment]::NewLine)
                
                Add-ColoredText -RichTextBox $detailsTextBox -Text "Bureau: " -Color $primaryColor -Style Bold
                Add-ColoredText -RichTextBox $detailsTextBox -Text ([string]$userDetails.PhysicalDeliveryOfficeName + [Environment]::NewLine)
                
                Add-ColoredText -RichTextBox $detailsTextBox -Text "Ville: " -Color $primaryColor -Style Bold
                Add-ColoredText -RichTextBox $detailsTextBox -Text ([string]$userDetails.City + [Environment]::NewLine)
                
                Add-ColoredText -RichTextBox $detailsTextBox -Text "Pays: " -Color $primaryColor -Style Bold
                Add-ColoredText -RichTextBox $detailsTextBox -Text ([string]$userDetails.Country + [Environment]::NewLine + [Environment]::NewLine)
                
                # Recuperer et afficher les groupes de l'utilisateur
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

# Affichage de la fenetre
[void]$form.ShowDialog()
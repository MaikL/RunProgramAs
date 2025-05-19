using namespace System.IO

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Storage locations
$scriptDataPath = "$env:APPDATA\RunProgramAs"
$credentialPath = Join-Path $scriptDataPath "credentials.xml"
$settingsPath = Join-Path $scriptDataPath "settings.json"
# Define log directory and file
$logDirectory = Join-Path -Path $PSScriptRoot -ChildPath "Log"
$logFileName = "RunProgramAs.log"
$logFilePath = Join-Path -Path $logDirectory -ChildPath $logFileName
$global:translations = $null
$global:currentLanguage = $null
$global:creds = @()
# Create directory for Credential data and Settings if it does not exist
if (-not (Test-Path $scriptDataPath)) {
    New-Item -ItemType Directory -Path $scriptDataPath | Out-Null
}
# Ensure the log directory exists
if (-not (Test-Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory | Out-Null
}
# Write-Log function
function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet("DEBUG", "INFO", "WARN", "ERROR")]
        [string]$Level = "INFO"
    )
    if ($Level -ge $global:logLevel) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "$timestamp [$Level] $Message"

        try {
            Set-LogRotation
            Add-Content -Path $logFilePath -Value $logEntry
        }
        catch {
            Write-Error "Failed to write to log file: $_"
        }
    }
}
# Function to rotate log if too big (> 1 MB)
function Set-LogRotation {
    $maxSizeMB = 1
    if (Test-Path $logFilePath) {
        $fileSizeMB = (Get-Item $logFilePath).Length / 1MB
        if ($fileSizeMB -ge $maxSizeMB) {
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $archivedLog = Join-Path -Path $logDirectory -ChildPath "RunProgramAs_$timestamp.log"
            Rename-Item -Path $logFilePath -NewName $archivedLog
        }
    }
}
# gets the global translations variable
# and sets it to the global variable translations
# This is a workaround to avoid using global variables directly in the script
function Get-TranslationForValue {
    param (
        [Parameter(Mandatory = $true)]
        [string]$value
    )
    return $global:translations.$global:currentLanguage.$value
}
# Function to test directory permissions
function Test-DirectoryPermissions {
    param ([string]$path)
    $user = $global:creds.UserName;
    try {
        $acl = Get-Acl $path
        foreach ($access in $acl.Access) {
            if ($access.IdentityReference.Value -like "*$user") {
                if ($access.FileSystemRights.ToString().Contains("Read") -or $access.FileSystemRights.ToString().Contains("Modify")) {
                    Write-Log "User: {$user} can read {$path}"
                    return $true
                }
                else {
                    Write-Log($access.FileSystemRights.ToString())
                }
            }
        }

        Write-Log "Test-DirectoryPermissions for {$user}"
        return $false
    }
    catch {
        Write-Log "Error in Test-DirectoryPermissions"
        return $false
    }
}
# Function to save multiple credentials
function Save-Credentials {
    param()

    $newCred = Get-Credential
    $credentialName = $newCred.UserName   # Use username as credential name

    $credentialStore = @{}
    if (Test-Path $credentialPath) {
        $credentialStore = Import-Clixml -Path $credentialPath
    }

    $credentialStore[$credentialName] = $newCred
    $credentialStore | Export-Clixml -Path $credentialPath

    return $credentialStore
}

# Function to load stored credentials
function Get-CredentialsFromJSON {
    if (Test-Path $credentialPath) {
        return Import-Clixml -Path $credentialPath
    }
    else {
        return @{}
    }
}
# Function to save settings
function Save-Settings($lastCommand, $arguments) {
    $settings = @{
        LastCommand = $lastCommand
        Arguments   = $arguments
    }
    $settings | ConvertTo-Json | Set-Content -Path $settingsPath
}
# Function to load settings
function Get-SettingsFromJSON {
    if (Test-Path $settingsPath) {
        $json = Get-Content $settingsPath -Raw | ConvertFrom-Json
        return $json
    }
    else {
        return [PSCustomObject]@{
            LastCommand = ""
            Arguments   = ""
        }
    }
}
# Function to update UI elements
function Update-UI {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$guiElements,
        [Parameter(Mandatory = $true)]
        [string] $currentLanguage
    )
    Write-Log "Update-UI Language: $currentLanguage"
    # Update global variables
    $global:currentLanguage = $currentLanguage
    $translations = $global:translations

    $guiElements.form.Text = Get-TranslationForValue -value "programName"
    $guiElements.lblFile.Text = Get-TranslationForValue -value filePath

    # Button Texte aktualisieren
    $guiElements.btnStart.Text = Get-TranslationForValue -value startProgram
    $guiElements.btnNewCreds.Text = Get-TranslationForValue -value newCredentials
    $guiElements.lblSavedCredentials.Text = Get-TranslationForValue -value savedCredentials
    $guiElements.lblForUser.Text = Get-TranslationForValue -value selectedCredential
    $guiElements.btnExit.Text = Get-TranslationForValue -value "exit"
    $guiElements.menuItemDelete.Text = Get-TranslationForValue -value delete

    # ComboBox und Labels aktualisieren
    $guiElements.lblFile.Text = Get-TranslationForValue -value filePath
    $guiElements.lblArgs.Text = Get-TranslationForValue -value arguments
    $guiElements.lblStatus.Text = ""

    # Aktualisiere alle anderen UI-Elemente, die textbasiert sind
    $guiElements.cmbCredentials.Items.Clear()
    $guiElements.cmbCredentials.Items.AddRange($global:credentialStore.Keys)
    if ($guiElements.cmbCredentials.Items.Count -gt 0) {
        $guiElements.cmbCredentials.SelectedIndex = 0
        $global:creds = $global:credentialStore[$guiElements.cmbCredentials.SelectedItem]
        $guiElements.lblUsername.Text = $global:creds.UserName
    }
}
# add buttons for each language
# This function creates buttons for each language in the translations file
function Add-LanguageButtons {
    param (
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.Panel]$panel,
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$GlobalVars
    )

    $yPos = 330
    $xPos = 30
    $languageButtons = @{}  # Dictionary, um Buttons mit Sprachen zu speichern

    foreach ($language in $GlobalVars.Translations.PSObject.Properties.Name) {
        $button = New-Object System.Windows.Forms.Button
        $button.Text = $GlobalVars.Translations.$language.Language
        $button.Location = New-Object System.Drawing.Point($xPos, $yPos)
        $button.Size = New-Object System.Drawing.Size(90, 25)
        $button.Tag = $language
        $panel.Controls.Add($button)

        # Speichere den Button im Dictionary
        $languageButtons[$language] = $button

        $xPos += 110
        if ($xPos -gt 300) {
            $xPos = 30
            $yPos += 30
        }
    }

    return @{
        Buttons = $languageButtons
        LastYPos = $yPos
    }
}

# Function to set global translations variable
# This function loads the translations from the JSON file and sets it to the global variable
function Set-GlobalTranslations {
    $translationsPath = "translations.json"
    $global:translations = Get-Content -Raw -Encoding utf8 -Path $translationsPath | ConvertFrom-Json
}
# Function to initialize language
# This function checks the current system locale and sets the language accordingly
function Initialize-Language {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$guiElements
    )
    Write-Log "Initializing language..."

    $locale = [System.Globalization.CultureInfo]::CurrentCulture
    if ($supportedLanguages -contains $locale.TwoLetterISOLanguageName) {
        $currentLanguage = $locale.TwoLetterISOLanguageName
        Write-Log "Matched locale: $currentLanguage"
    }
    else {
        $currentLanguage = "en"  # Default to English
        Write-Log "No match found. Defaulting to English."
    }

    Update-UI -currentLanguage $currentLanguage -guiElements $guiElements
}

# Function to select a command file
# This function opens a file dialog to select a command file
function Select-Command {
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = Get-TranslationForValue -value selectFile
    $openFileDialog.Filter = Get-TranslationForValue -value programFilter
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')

    if ($openFileDialog.ShowDialog() -eq "OK") {
        return $openFileDialog.FileName
    }
    else {
        return $null
    }
}
# Function to update the ComboBox with saved credentials
# This function is called when the form is loaded or when new credentials are added
function Update-CredentialComboBox {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$guiElements
    )

    $cmbCredentials = $guiElements.CmbCredentials
    $cmbCredentials.Items.Clear()
    $cmbCredentials.Items.AddRange($global:credentialStore.Keys)
    if ($cmbCredentials.Items.Count -gt 0) {
        $cmbCredentials.SelectedIndex = 0
        $global:creds = $global:credentialStore[$cmbCredentials.SelectedItem]
    }
    else {
        $global:creds = $null
    }
}




# Executes the function to set global translations variable
# This function is called at the beginning of the script to load translations
Set-GlobalTranslations
# Getting all translations dynamically from translations.json
$supportedLanguages = $global:translations.PSObject.Properties.Name

# Load settings and credentials
$global:settings = Get-SettingsFromJSON
$global:credentialStore = Get-CredentialsFromJSON
$global:creds = $null

function Get-RunProgramAsForm {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$globalVars
    )
    # Create GUI
    $form = New-Object System.Windows.Forms.Form
    $form.Text = Get-TranslationForValue -value programName
    $form.Size = New-Object System.Drawing.Size(500, 500)
    $form.StartPosition = "CenterScreen"

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Location = New-Object System.Drawing.Point(0, 0)
    $panel.Size = New-Object System.Drawing.Size($form.ClientSize.Width, $form.ClientSize.Height)  # dynamic size
    $panel.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $panel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
    $panel.AutoScroll = $true
    $form.Controls.Add($panel)

    # Button: Start program
    $btnStart = New-Object System.Windows.Forms.Button
    $btnStart.Location = New-Object System.Drawing.Point(30, 30)
    $btnStart.Size = New-Object System.Drawing.Size(200, 25)
    $btnStart.Text = Get-TranslationForValue -value startProgram

    # Button: New credentials
    $btnNewCreds = New-Object System.Windows.Forms.Button
    $btnNewCreds.Location = New-Object System.Drawing.Point(250, 30)
    $btnNewCreds.Size = New-Object System.Drawing.Size(200, 25)
    $btnNewCreds.Text = Get-TranslationForValue -value newCredentials

    # Label: saved credentials
    $lblSavedCredentials = New-Object System.Windows.Forms.Label
    $lblSavedCredentials.Location = New-Object System.Drawing.Point(30, 70)
    $lblSavedCredentials.Size = New-Object System.Drawing.Size(250, 20)
    $lblSavedCredentials.Text = Get-TranslationForValue -value savedCredentials

    # ComboBox: Selection of saved credentials
    $cmbCredentials = New-Object System.Windows.Forms.ComboBox
    $cmbCredentials.Location = New-Object System.Drawing.Point(30, 90)
    $cmbCredentials.Size = New-Object System.Drawing.Size(420, 21)
    $cmbCredentials.DropDownStyle = 'DropDownList'
    $cmbCredentials.IntegralHeight = $false
    $cmbCredentials.MaxDropDownItems = 10

    # Label: selected user
    $lblForUser = New-Object System.Windows.Forms.Label
    $lblForUser.Location = New-Object System.Drawing.Point(30, 120)
    $lblForUser.Size = New-Object System.Drawing.Size(130, 20)
    $lblForUser.Text = Get-TranslationForValue -value selectedCredential
    # Label: selected user
    $lblUsername = New-Object System.Windows.Forms.Label
    $lblUsername.Location = New-Object System.Drawing.Point(160, 120)
    $lblUsername.Size = New-Object System.Drawing.Size(200, 20)
    $lblUsername.Text = $global:creds.UserName

    # Label: file
    $lblFile = New-Object System.Windows.Forms.Label
    $lblFile.Location = New-Object System.Drawing.Point(30, 140)
    $lblFile.Size = New-Object System.Drawing.Size(380, 20)
    $lblFile.Text = Get-TranslationForValue -value filePath

    # last used command
    $txtFile = New-Object System.Windows.Forms.TextBox
    $txtFile.Location = New-Object System.Drawing.Point(30, 160)
    $txtFile.Size = New-Object System.Drawing.Size(380, 20)
    $txtFile.Text = $global:settings.LastCommand

    # browse for new command
    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Location = New-Object System.Drawing.Point(410, 158)
    $btnBrowse.Size = New-Object System.Drawing.Size(40, 24)
    $btnBrowse.Text = "..."

    # label: arguments
    $lblArgs = New-Object System.Windows.Forms.Label
    $lblArgs.Location = New-Object System.Drawing.Point(30, 185)
    $lblArgs.Size = New-Object System.Drawing.Size(380, 20)
    $lblArgs.Text = Get-TranslationForValue -value arguments

    # textbox for arguments
    $txtArgs = New-Object System.Windows.Forms.TextBox
    $txtArgs.Location = New-Object System.Drawing.Point(30, 210)
    $txtArgs.Size = New-Object System.Drawing.Size(420, 20)
    $txtArgs.Text = $global:settings.Arguments

    # Status-Label
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(30, 230)
    $lblStatus.Size = New-Object System.Drawing.Size(420, 50)
    $lblStatus.Text = ""

    # Button: Exit
    $btnExit = New-Object System.Windows.Forms.Button
    $btnExit.Location = New-Object System.Drawing.Point(140, 290)
    $btnExit.Size = New-Object System.Drawing.Size(150, 25)
    $btnExit.Text = Get-TranslationForValue -value "exit"

    Write-Log "lastButtonPosition: $lastButtonPosition" "DEBUG"
    # LinkLabel: Github Page
    $linkGitHub = New-Object System.Windows.Forms.LinkLabel
    $linkGitHub.Size = New-Object System.Drawing.Size(420, 20)
    $linkGitHub.Text = Get-TranslationForValue -value visitGitHub
    $linkGitHub.Location = New-Object System.Drawing.Point(30, (400 + 40))

    $linkGitHub.add_LinkClicked({
            try {
                Start-Process "https://github.com/MaikL/RunProgramAs"
            }
            catch {
                Write-Log "Fehler beim Öffnen des Links: $_" "ERROR"
            }
        })
    # ContextMenu for ComboBox
    $contextMenu = New-Object System.Windows.Forms.ContextMenu
    $menuItemDelete = New-Object System.Windows.Forms.MenuItem (Get-TranslationForValue -value delete)

    $panel.Controls.Add($btnStart)
    $panel.Controls.Add($btnNewCreds)
    $panel.Controls.Add($lblSavedCredentials)
    $panel.Controls.Add($cmbCredentials)
    $panel.Controls.Add($lblForUser)
    $panel.Controls.Add($lblUsername)
    $panel.Controls.Add($lblFile)
    $panel.Controls.Add($txtFile)
    $panel.Controls.Add($btnBrowse)
    $panel.Controls.Add($lblArgs)
    $panel.Controls.Add($txtArgs)
    $panel.Controls.Add($lblStatus)
    $panel.Controls.Add($btnExit)
    $panel.Controls.Add($linkGitHub)

    $form.Height = $linkGitHub.Location.Y + $linkGitHub.Size.Height + 50

    return @{
        form                = $form
        btnStart            = $btnStart
        btnNewCreds         = $btnNewCreds
        cmbCredentials      = $cmbCredentials
        lblSavedCredentials = $lblSavedCredentials
        lblForUser          = $lblForUser
        lblUsername         = $lblUsername
        lblFile             = $lblFile
        txtFile             = $txtFile
        btnBrowse           = $btnBrowse
        lblArgs             = $lblArgs
        txtArgs             = $txtArgs
        lblStatus           = $lblStatus
        btnExit             = $btnExit
        linkGitHub          = $linkGitHub
        contextMenu         = $contextMenu
        menuItemDelete      = $menuItemDelete
        panel               = $panel
    }
}

function Add-Events {
    [CmdletBinding()]
    param (
        $guiElements
    )
    # Events
    $guiElements.btnBrowse.Add_Click({
            $command = Select-Command
            if ($command) {
                $guiElements.txtFile.Text = $command
            }
        })

    $guiElements.btnStart.Add_Click({
            $command = $guiElements.txtFile.Text.Trim()
            $arguments = $guiElements.txtArgs.Text.Trim()

            # check if a command is given
            if ([string]::IsNullOrWhiteSpace($command)) {
                [System.Windows.Forms.MessageBox]::Show($global:translations.$currentLanguage.noPathSelected, $global:translations.$currentLanguage.error, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $command = Select-Command
                if ($command) {
                    $guiElements.txtFile.Text = $command
                }
                else {
                    $guiElements.lblStatus.ForeColor = 'Orange'
                    $guiElements.lblStatus.Text = Get-TranslationForValue -value cancelNoFile
                    return
                }
            }

            # check: file exists?
            if (-not [System.IO.File]::Exists($command)) {
                $guiElements.lblStatus.ForeColor = 'Red'
                $guiElements.lblStatus.Text = Get-TranslationForValue -value FileNotFound
                return
            }

            # check is executable?
            $validExtensions = @(".exe", ".bat", ".cmd")
            if (-not ($validExtensions -contains [System.IO.Path]::GetExtension($command).ToLower())) {
                $guiElements.lblStatus.ForeColor = 'Red'
                $guiElements.lblStatus.Text = Get-TranslationForValue -value onlyExecutableFiles
                return
            }

            # Save settings
            Save-Settings -lastCommand $command -arguments $arguments
            $commandPath = Split-Path $command -Parent
            if (-not (Test-DirectoryPermissions -path $commandPath)) {
                Write-Log "Error: User has no permission for path: $commandPath" "ERROR"
                $guiElements.lblStatus.ForeColor = 'Red'
                $guiElements.lblStatus.Text = Get-TranslationForValue -value noReadingPermission
                return
            }

            # Debugging: check username
            Write-Log "Username = $($global:creds.UserName)" "DEBUG"
            # starting process, depending if the Path has arguments or not
            try {
                if ([string]::IsNullOrWhiteSpace($arguments)) {
                    Start-Process -FilePath $command `
                        -WorkingDirectory (Split-Path $command -Parent) `
                        -Credential $global:creds `
                        -ErrorAction Stop
                }
                else {
                    Start-Process -FilePath $command `
                        -WorkingDirectory (Split-Path $command -Parent) `
                        -Credential $global:creds `
                        -ArgumentList $arguments `
                        -ErrorAction Stop
                }
                $guiElements.lblStatus.ForeColor = 'Green'
                $guiElements.lblStatus.Text = Get-TranslationForValue -value programStartSuccessful
            }
            catch {
                Write-Log "starting process $command error: $_" "ERROR"

                $guiElements.lblStatus.ForeColor = 'Red'
                $guiElements.lblStatus.Text = "{$global:translations.$currentLanguage.startingError}`n$($_.Exception.Message)"
            }
        })

    # add new credentials
    $guiElements.btnNewCreds.Add_Click({
            try {
                $global:credentialStore = Save-Credentials
                $guiElements.lblStatus.ForeColor = 'Blue'
                $guiElements.lblStatus.Text = Get-TranslationForValue -value savedNewCredentials

                # update ComboBox
                $guiElements.cmbCredentials.Items.Clear()
                Update-CredentialComboBox -guiElements $guiElements
                $guiElements.cmbCredentials.SelectedItem = ($global:credentialStore.Keys | Select-Object -Last 1)
                $global:creds = $global:credentialStore[$guiElements.cmbCredentials.SelectedItem]
            }
            finally {
                $guiElements.form.Enabled = $true  # reactivate form
                $guiElements.form.Activate()       # bring form to foreground
            }
        })
    # exit button action
    $guiElements.btnExit.Add_Click({
            $guiElements.form.Close()
        })
    # action when select changed
    $guiElements.cmbCredentials.Add_SelectedIndexChanged({
            $selectedKey = $guiElements.cmbCredentials.SelectedItem
            if ($selectedKey) {
                $global:creds = $global:credentialStore[$selectedKey]
                $guiElements.lblUsername.Text = $selectedKey
            }
        })

    # delete selected credentials
    $guiElements.menuItemDelete.add_Click({
            $selectedKey = $guiElements.cmbCredentials.SelectedItem
            if ($selectedKey) {
                $txtDelete = $global:translations.$global:currentLanguage.confirmDelete -replace "{selectedKey}", $selectedKey

                if ([System.Windows.Forms.MessageBox]::Show("$txtDelete", $global:translations.$global:currentLanguage.confirmation, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning) -eq [System.Windows.Forms.DialogResult]::Yes) {
                    $global:credentialStore.Remove($selectedKey)
                    $global:credentialStore | Export-Clixml -Path $credentialPath

                    # update ComboBox
                    $guiElements.cmbCredentials.Items.Clear()
                    $guiElements.cmbCredentials.Items.AddRange($global:credentialStore.Keys)

                    if ($guiElements.cmbCredentials.Items.Count -gt 0) {
                        $guiElements.cmbCredentials.SelectedIndex = 0
                        $global:creds = $global:credentialStore[$guiElements.cmbCredentials.SelectedItem]
                    }
                    else {
                        $global:creds = $null
                    }

                    $guiElements.lblStatus.ForeColor = 'Blue'
                    $guiElements.lblStatus.Text = Get-TranslationForValue -value credentialsDeleted
                }
            }
        })
}
# Globale Variablen vorbereiten
$globalVars = [PSCustomObject]@{
    Translations    = $global:translations
    CurrentLanguage = $global:currentLanguage
    CredentialStore = $global:credentialStore
    Settings        = $global:settings
}


Write-Log "Starting RunProgramAs GUI" "INFO"
$guiElements = Get-RunProgramAsForm -globalVars $globalVars
# gets the translations variable and sets it to the global variable translations
Write-Log "Initialize Language" "INFO"
Initialize-Language -guiElements $guiElements

Write-Log "Adding Events" "INFO"
Add-Events -guiElements $guiElements

$languageButtonResult = Add-LanguageButtons -Panel $guiElements.panel -GlobalVars $globalVars
$languageButtons = $languageButtonResult.Buttons
$lastButtonPosition = $languageButtonResult.LastYPos

# Click-Events nachträglich hinzufügen
foreach ($language in $languageButtons.Keys) {
    $button = $languageButtons[$language]
    $button.Add_Click({
        $guiElements.CurrentLanguage = $this.Tag
        Write-Log "Language button clicked: $($this.Tag)"
        Update-UI -currentLanguage $this.Tag -guiElements $guiElements
    })
}

Write-Log "Adding Menu" "INFO"
$guiElements.contextMenu.MenuItems.Add($guiElements.menuItemDelete)
$guiElements.cmbCredentials.ContextMenu = $guiElements.contextMenu
Update-CredentialComboBox -guiElements $guiElements
# show GUI
$guiElements.Form.Topmost = $true
[void]$guiElements.Form.ShowDialog()

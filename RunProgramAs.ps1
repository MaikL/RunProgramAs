using namespace System.IO

param (
    [Parameter(Mandatory=$false)]
    [SecureString]$Credentials
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Storage locations
$scriptDataPath = "$env:APPDATA\RunProgramAs"
$credentialPath = Join-Path $scriptDataPath "credentials.xml"
$settingsPath = Join-Path $scriptDataPath "settings.json"
$global:translations = @()
$global:creds = @()
# Create directory if it does not exist
if (-not (Test-Path $scriptDataPath)) {
    New-Item -ItemType Directory -Path $scriptDataPath | Out-Null
}

# Define log directory and file
$logDirectory = Join-Path -Path $PSScriptRoot -ChildPath "Log"
$logFileName = "RunProgramAs.log"
$logFilePath = Join-Path -Path $logDirectory -ChildPath $logFileName

# Ensure the log directory exists
if (-not (Test-Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory | Out-Null
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

# Write-Log function
function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet("DEBUG","INFO", "WARN", "ERROR")]
        [string]$Level = "INFO"
    )

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

# Function to load translations
function Get-Translations {
    $translationsPath = "translations.json"
    $global:translations = Get-Content -Raw -Encoding utf8 -Path $translationsPath | ConvertFrom-Json
}

# Function to update UI elements
function Update-UI {
    param (
        [string] $currentLanguage
    )
    Write-Log "Update-UI Language: $currentLanguage"
    $translations = $global:translations

    $form.Text = $translations.$currentLanguage.programName
    $lblFile.Text = $translations.$currentLanguage.filePath

    # Button Texte aktualisieren
    $btnStart.Text = $translations.$currentLanguage.startProgram
    $btnNewCreds.Text = $translations.$currentLanguage.newCredentials
    $lblSavedCredentials.Text = $global:translations.$currentLanguage.savedCredentials
    $lblForUser.Text = $translations.$currentLanguage.selectedCredential
    $btnExit.Text = $translations.$currentLanguage.exit
    $menuItemDelete.Text = $global:translations.$currentLanguage.delete

    $btnDeutsch.Text = "Deutsch"
    $btnEnglisch.Text = "English"

    # ComboBox und Labels aktualisieren
    $lblFile.Text = $translations.$currentLanguage.filePath
    $lblArgs.Text = $translations.$currentLanguage.arguments
    $lblStatus.Text = ""

    # Aktualisiere alle anderen UI-Elemente, die textbasiert sind
    $cmbCredentials.Items.Clear()
    $cmbCredentials.Items.AddRange($global:credentialStore.Keys)
    if ($cmbCredentials.Items.Count -gt 0) {
        $cmbCredentials.SelectedIndex = 0
        $global:creds = $global:credentialStore[$cmbCredentials.SelectedItem]
        $lblUsername.Text = $global:creds.UserName
    }
}

Get-Translations
$locale = Get-WinSystemLocale
Write-Log "local system language $locale.TwoLetterISOLanguageName"
# Getting all translations dynamically from translations.json
$supportedLanguages = $global:translations.PSObject.Properties.Name

if ($supportedLanguages -match $locale.TwoLetterISOLanguageName) {
    $currentLanguage = $locale.TwoLetterISOLanguageName
    Write-Log "matched locale $currentLanguage"
}
else {
    $currentLanguage = "en"  # English as standard language
}
function Initialize-Language {
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

    Update-UI -currentLanguage $currentLanguage
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
# Function to save settings
function Save-Settings($lastCommand, $arguments) {
    $settings = @{
        LastCommand = $lastCommand
        Arguments   = $arguments
    }
    $settings | ConvertTo-Json | Set-Content -Path $settingsPath
}

# Function to select a command
function Select-Command {
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = $global:translations.$currentLanguage.selectFile
    $openFileDialog.Filter = $global:translations.$currentLanguage.programFilter
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')

    if ($openFileDialog.ShowDialog() -eq "OK") {
        return $openFileDialog.FileName
    }
    else {
        return $null
    }
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

# Load settings and credentials
$global:settings = Get-SettingsFromJSON
$global:credentialStore = Get-CredentialsFromJSON
$global:creds = $null

# Create GUI
$form = New-Object System.Windows.Forms.Form
$form.Text = $global:translations.$currentLanguage.programName
$form.Size = New-Object System.Drawing.Size(450, 450)
$form.StartPosition = "CenterScreen"

# Button: Start program
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Location = New-Object System.Drawing.Point(30, 30)
$btnStart.Size = New-Object System.Drawing.Size(180, 25)
$btnStart.Text = $global:translations.$currentLanguage.startProgram
$form.Controls.Add($btnStart)

# Button: New credentials
$btnNewCreds = New-Object System.Windows.Forms.Button
$btnNewCreds.Location = New-Object System.Drawing.Point(220, 30)
$btnNewCreds.Size = New-Object System.Drawing.Size(180, 25)
$btnNewCreds.Text = $global:translations.$currentLanguage.newCredentials
$form.Controls.Add($btnNewCreds)

# Label: saved credentials
$lblSavedCredentials = New-Object System.Windows.Forms.Label
$lblSavedCredentials.Location = New-Object System.Drawing.Point(30, 70)
$lblSavedCredentials.Size = New-Object System.Drawing.Size(370, 20)
$lblSavedCredentials.Text = $global:translations.$currentLanguage.savedCredentials
$form.Controls.Add($lblSavedCredentials)

# ComboBox: Selection of saved credentials
$cmbCredentials = New-Object System.Windows.Forms.ComboBox
$cmbCredentials.Location = New-Object System.Drawing.Point(30, 90)
$cmbCredentials.Size = New-Object System.Drawing.Size(370, 21)
$cmbCredentials.DropDownStyle = 'DropDownList'
# Load saved credentials and display usernames
$cmbCredentials.Items.AddRange($global:credentialStore.Keys)

# If any exist, select the first one
if ($cmbCredentials.Items.Count -gt 0) {
    $cmbCredentials.SelectedIndex = 0
    $global:creds = $global:credentialStore[$cmbCredentials.SelectedItem]
}

$form.Controls.Add($cmbCredentials)

# Label: selected user
$lblForUser = New-Object System.Windows.Forms.Label
$lblForUser.Location = New-Object System.Drawing.Point(30, 120)
$lblForUser.Size = New-Object System.Drawing.Size(110, 20)
$lblForUser.Text = $global:translations.$currentLanguage.selectedCredential
$form.Controls.Add($lblForUser)
# Label: selected user
$lblUsername = New-Object System.Windows.Forms.Label
$lblUsername.Location = New-Object System.Drawing.Point(150, 120)
$lblUsername.Size = New-Object System.Drawing.Size(200, 20)
$lblUsername.Text = $global:creds.UserName
$form.Controls.Add($lblUsername)

# Label: file
$lblFile = New-Object System.Windows.Forms.Label
$lblFile.Location = New-Object System.Drawing.Point(30, 140)
$lblFile.Size = New-Object System.Drawing.Size(380, 20)
$lblFile.Text = $global:translations.$currentLanguage.filePath
$form.Controls.Add($lblFile)

# last used command
$txtFile = New-Object System.Windows.Forms.TextBox
$txtFile.Location = New-Object System.Drawing.Point(30, 160)
$txtFile.Size = New-Object System.Drawing.Size(320, 20)
$txtFile.Text = $global:settings.LastCommand
$form.Controls.Add($txtFile)

# browse for new command
$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Location = New-Object System.Drawing.Point(360, 158)
$btnBrowse.Size = New-Object System.Drawing.Size(40, 24)
$btnBrowse.Text = "..."
$form.Controls.Add($btnBrowse)

# label: arguments
$lblArgs = New-Object System.Windows.Forms.Label
$lblArgs.Location = New-Object System.Drawing.Point(30, 185)
$lblArgs.Size = New-Object System.Drawing.Size(380, 20)
$lblArgs.Text = $global:translations.$currentLanguage.arguments
$form.Controls.Add($lblArgs)

# textbox for arguments
$txtArgs = New-Object System.Windows.Forms.TextBox
$txtArgs.Location = New-Object System.Drawing.Point(30, 210)
$txtArgs.Size = New-Object System.Drawing.Size(370, 20)
$txtArgs.Text = $global:settings.Arguments
$form.Controls.Add($txtArgs)

# Status-Label
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(30, 230)
$lblStatus.Size = New-Object System.Drawing.Size(380, 50)
$lblStatus.Text = ""
$form.Controls.Add($lblStatus)

# Button: Exit
$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Location = New-Object System.Drawing.Point(140, 290)
$btnExit.Size = New-Object System.Drawing.Size(150, 25)
$btnExit.Text = $global:translations.$currentLanguage.exit
$form.Controls.Add($btnExit)


# Buttons for German and English
$btnDeutsch = New-Object System.Windows.Forms.Button
$btnDeutsch.Text = "Deutsch"
$btnDeutsch.Size = New-Object System.Drawing.Size(100, 25)
$btnDeutsch.Location = New-Object System.Drawing.Point(100, 330)
$form.Controls.Add($btnDeutsch)

$btnEnglisch = New-Object System.Windows.Forms.Button
$btnEnglisch.Text = "English"
$btnEnglisch.Size = New-Object System.Drawing.Size(100, 25)
$btnEnglisch.Location = New-Object System.Drawing.Point(250, 330)

$form.Controls.Add($btnEnglisch)

# Eventhandler for language buttons
$btnDeutsch.Add_Click({
        $global:currentLanguage = "de"
        $currentLanguage = "de"
        Get-Translations
        Update-UI -currentLanguage "de"
    })
    $btnEnglisch.Add_Click({
        $global:currentLanguage = "en"
        $currentLanguage = "en"
        Get-Translations
        Update-UI -currentLanguage "en"
    })

# Events
$btnBrowse.Add_Click({
        $command = Select-Command
        if ($command) {
            $txtFile.Text = $command
        }
    })

$btnStart.Add_Click({
        $command = $txtFile.Text.Trim()
        $arguments = $txtArgs.Text.Trim()

        # check if a command is given
        if ([string]::IsNullOrWhiteSpace($command)) {
            [System.Windows.Forms.MessageBox]::Show($global:translations.$currentLanguage.noPathSelected, $global:translations.$currentLanguage.error, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $command = Select-Command
            if ($command) {
                $txtFile.Text = $command
            }
            else {
                $lblStatus.ForeColor = 'Orange'
                $lblStatus.Text = $global:translations.$currentLanguage.cancelNoFile
                return
            }
        }

        # check: file exists?
        if (-not [System.IO.File]::Exists($command)) {
            $lblStatus.ForeColor = 'Red'
            $lblStatus.Text = $global:translations.$currentLanguage.FileNotFound
            return
        }

        # check is executable?
        $validExtensions = @(".exe", ".bat", ".cmd")
        if (-not ($validExtensions -contains [System.IO.Path]::GetExtension($command).ToLower())) {
            $lblStatus.ForeColor = 'Red'
            $lblStatus.Text = $global:translations.$currentLanguage.onlyExecutableFiles
            return
        }

        # Save settings
        Save-Settings -lastCommand $command -arguments $arguments
        $commandPath = Split-Path $command -Parent
        if (-not (Test-DirectoryPermissions -path $commandPath)) {
            Write-Log "Error: User has no permission for path: $commandPath" "ERROR"
            $lblStatus.ForeColor = 'Red'
            $lblStatus.Text = $global:translations.$currentLanguage.noReadingPermission
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
            $lblStatus.ForeColor = 'Green'
            $lblStatus.Text = $global:translations.$currentLanguage.programStartSuccessful
        }
        catch {
            Write-Log "starting process $command error: $_" "ERROR"

            $lblStatus.ForeColor = 'Red'
            $lblStatus.Text = "{$global:translations.$currentLanguage.startingError}`n$($_.Exception.Message)"
        }
    })

# add new credentials
$btnNewCreds.Add_Click({
        $form.Enabled = $false  # disable form
        try {
            $global:credentialStore = Save-Credentials
            $lblStatus.ForeColor = 'Blue'
            $lblStatus.Text = $global:translations.$currentLanguage.savedNewCredentials

            # update ComboBox
            $cmbCredentials.Items.Clear()
            $cmbCredentials.Items.AddRange($global:credentialStore.Keys)
            $cmbCredentials.SelectedItem = ($global:credentialStore.Keys | Select-Object -Last 1)
            $l
        }
        finally {
            $form.Enabled = $true  # reactivate form
            $form.Activate()       # bring form to foreground
        }
    })
# exit button action
$btnExit.Add_Click({
        $form.Close()
    })
# action when select changed
$cmbCredentials.Add_SelectedIndexChanged({
        $selectedKey = $cmbCredentials.SelectedItem
        if ($selectedKey) {
            $global:creds = $global:credentialStore[$selectedKey]
            $lblUsername.Text = $selectedKey
        }
    })

# ContextMenu for ComboBox
$contextMenu = New-Object System.Windows.Forms.ContextMenu
$menuItemDelete = New-Object System.Windows.Forms.MenuItem $global:translations.$currentLanguage.delete

# delete selected credentials
$menuItemDelete.add_Click({
        $selectedKey = $cmbCredentials.SelectedItem
        if ($selectedKey) {
            $txtDelete = $global:translations.$global:currentLanguage.confirmDelete.Replace("selectedKey", $selectedKey)
            if ([System.Windows.Forms.MessageBox]::Show("$txtDelete", $global:translations.$global:currentLanguage.confirmation, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning) -eq [System.Windows.Forms.DialogResult]::Yes) {
                $global:credentialStore.Remove($selectedKey)
                $global:credentialStore | Export-Clixml -Path $credentialPath

                # update ComboBox
                $cmbCredentials.Items.Clear()
                $cmbCredentials.Items.AddRange($global:credentialStore.Keys)

                if ($cmbCredentials.Items.Count -gt 0) {
                    $cmbCredentials.SelectedIndex = 0
                    $global:creds = $global:credentialStore[$cmbCredentials.SelectedItem]
                }
                else {
                    $global:creds = $null
                }

                $lblStatus.ForeColor = 'Blue'
                $lblStatus.Text = $global:translations.$currentLanguage.credentialsDeleted
            }
        }
    })

$contextMenu.MenuItems.Add($menuItemDelete)
$cmbCredentials.ContextMenu = $contextMenu
Initialize-Language
# show GUI
$form.Topmost = $true
[void]$form.ShowDialog()

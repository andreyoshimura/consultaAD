#requires -Modules ActiveDirectory
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Funções auxiliares ---
function Test-ADModule {
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        [System.Windows.Forms.MessageBox]::Show(
            "O módulo 'ActiveDirectory' não está instalado. Instale RSAT ou adicione a feature Active Directory Module for Windows PowerShell.",
            "Dependência ausente", 'OK', 'Error'
        ) | Out-Null
        return $false
    }
    return $true
}

function Find-ADUser([string]$Sam) {
    if ([string]::IsNullOrWhiteSpace($Sam)) { return $null }
    try {
        # Busca por sAMAccountName exatamente igual (ignore case já é padrão)
        return Get-ADUser -Filter "SamAccountName -eq '$Sam'"
    }
    catch {
        throw $_
    }
}

if (-not (Test-ADModule)) { return }

# --- UI ---
$form              = New-Object System.Windows.Forms.Form
$form.Text         = "Consulta de Usuário no AD"
$form.Size         = New-Object System.Drawing.Size(480,210)
$form.StartPosition= "CenterScreen"
$form.Topmost      = $false

$lblUser           = New-Object System.Windows.Forms.Label
$lblUser.Text      = "Usuário (sAMAccountName):"
$lblUser.Location  = New-Object System.Drawing.Point(20,25)
$lblUser.AutoSize  = $true

$txtUser           = New-Object System.Windows.Forms.TextBox
$txtUser.Location  = New-Object System.Drawing.Point(22,50)
$txtUser.Size      = New-Object System.Drawing.Size(420,24)

$btnCheck          = New-Object System.Windows.Forms.Button
$btnCheck.Text     = "Verificar"
$btnCheck.Location = New-Object System.Drawing.Point(22,85)
$btnCheck.Size     = New-Object System.Drawing.Size(100,32)

$lblResult         = New-Object System.Windows.Forms.Label
$lblResult.Text    = "Resultado aparecerá aqui."
$lblResult.Location= New-Object System.Drawing.Point(22,130)
$lblResult.AutoSize= $true
$lblResult.Font    = New-Object System.Drawing.Font("Segoe UI", 10,[System.Drawing.FontStyle]::Bold)

$form.Controls.AddRange(@($lblUser,$txtUser,$btnCheck,$lblResult))

# Permite pressionar Enter para acionar o botão
$form.AcceptButton = $btnCheck

# Lógica do botão
$btnCheck.Add_Click({
    $user = $txtUser.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($user)) {
        $lblResult.ForeColor = [System.Drawing.Color]::DarkRed
        $lblResult.Text = "Informe um nome de usuário."
        return
    }

    $lblResult.ForeColor = [System.Drawing.Color]::Black
    $lblResult.Text = "Consultando no AD..."

    try {
        $found = Find-ADUser -Sam $user
        if ($found) {
            $lblResult.ForeColor = [System.Drawing.Color]::ForestGreen
            $lblResult.Text = "✅ Usuário '$user' ENCONTRADO (DN: $($found.DistinguishedName))"
        } else {
            $lblResult.ForeColor = [System.Drawing.Color]::DarkRed
            $lblResult.Text = "❌ Usuário '$user' NÃO encontrado."
        }
    }
    catch {
        $lblResult.ForeColor = [System.Drawing.Color]::DarkRed
        $lblResult.Text = "Erro ao consultar o AD: $($_.Exception.Message)"
    }
})

[void]$form.ShowDialog()

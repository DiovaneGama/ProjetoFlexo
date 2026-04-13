# ============================================================
#  InstalarTarefa.ps1
#  Roda UMA VEZ em cada maquina, como Administrador.
#  Registra a tarefa agendada que executa AtualizarConsoleFlexo.ps1
#  silenciosamente a cada logon do Windows.
#
#  Como usar:
#    1. Copie este arquivo e AtualizarConsoleFlexo.ps1 para a maquina
#    2. Clique com botao direito > "Executar como Administrador"
#    3. Pronto -- a tarefa ficara registrada permanentemente
# ============================================================

# --- CONFIGURACAO ---
# Ajuste este caminho para onde o script de atualizacao ficara na maquina
$SCRIPT_LOCAL = "C:\ProgramData\ConsoleFlexo\AtualizarConsoleFlexo.ps1"
$NOME_TAREFA  = "ConsoleFlexo - Atualizacao Automatica"

# --- COPIA O SCRIPT DE ATUALIZACAO PARA LOCAL PERMANENTE ---
$pastaScript = Split-Path $SCRIPT_LOCAL
if (-not (Test-Path $pastaScript)) {
    New-Item -ItemType Directory -Path $pastaScript | Out-Null
}

$scriptOrigem = Join-Path $PSScriptRoot "AtualizarConsoleFlexo.ps1"
Copy-Item -Path $scriptOrigem -Destination $SCRIPT_LOCAL -Force
Write-Host "Script copiado para: $SCRIPT_LOCAL"

# --- CRIA A TAREFA AGENDADA ---
$acao = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$SCRIPT_LOCAL`""

# Dispara no logon de QUALQUER usuario da maquina
$gatilho = New-ScheduledTaskTrigger -AtLogOn

# Roda com privilegios do usuario logado, sem exigir senha
$configuracao = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 2) `
    -StartWhenAvailable `
    -DontStopIfGoingOnBatteries `
    -AllowStartIfOnBatteries

$principal = New-ScheduledTaskPrincipal `
    -GroupId "BUILTIN\Users" `
    -RunLevel Limited

# Remove tarefa anterior se existir
Unregister-ScheduledTask -TaskName $NOME_TAREFA -Confirm:$false -ErrorAction SilentlyContinue

Register-ScheduledTask `
    -TaskName  $NOME_TAREFA `
    -Action    $acao `
    -Trigger   $gatilho `
    -Settings  $configuracao `
    -Principal $principal `
    -Description "Atualiza ConsoleFlexo.gms da pasta de rede no logon do Windows."

Write-Host ""
Write-Host "Tarefa '$NOME_TAREFA' registrada com sucesso!" -ForegroundColor Green
Write-Host "A atualizacao ocorrera automaticamente a cada logon."
Write-Host ""
Write-Host "Para verificar: Agendador de Tarefas > Biblioteca > $NOME_TAREFA"

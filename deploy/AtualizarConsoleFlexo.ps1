# ============================================================
#  AtualizarConsoleFlexo.ps1
#  Executa no logon do Windows via Agendador de Tarefas.
#  Copia ConsoleFlexo.gms da pasta de rede para o GMS local
#  do CorelDRAW, de forma silenciosa e sem interacao do usuario.
# ============================================================

# --- CONFIGURACAO ---
$ORIGEM  = "\\servidor\ConsoleFlexo\ConsoleFlexo.gms"   # <-- ajuste o caminho de rede
$LOG     = "$env:APPDATA\ConsoleFlexo\atualizacao.log"
$COREL_PROCESSO = "CorelDRW"

# --- PASTA DE DESTINO: detecta automaticamente a versao instalada ---
$pastaCorel = Get-ChildItem "$env:APPDATA\Corel" -Directory -ErrorAction SilentlyContinue |
              Where-Object { $_.Name -like "CorelDRAW Graphics Suite*" } |
              Sort-Object Name -Descending |
              Select-Object -First 1

if (-not $pastaCorel) {
    Add-Content $LOG "$(Get-Date -Format 'yyyy-MM-dd HH:mm') | ERRO: CorelDRAW nao encontrado em AppData."
    exit 1
}

$DESTINO = Join-Path $pastaCorel.FullName "Draw\GMS\ConsoleFlexo.gms"

# --- GARANTE QUE A PASTA DE LOG EXISTE ---
$pastaLog = Split-Path $LOG
if (-not (Test-Path $pastaLog)) { New-Item -ItemType Directory -Path $pastaLog | Out-Null }

# --- VERIFICA SE O COREL ESTA ABERTO ---
if (Get-Process -Name $COREL_PROCESSO -ErrorAction SilentlyContinue) {
    Add-Content $LOG "$(Get-Date -Format 'yyyy-MM-dd HH:mm') | AVISO: CorelDRAW esta aberto. Atualizacao adiada para o proximo logon."
    exit 0
}

# --- VERIFICA SE A REDE ESTA DISPONIVEL ---
if (-not (Test-Path $ORIGEM)) {
    Add-Content $LOG "$(Get-Date -Format 'yyyy-MM-dd HH:mm') | AVISO: Pasta de rede inacessivel. Sem atualizacao."
    exit 0
}

# --- COMPARA SE JA E A MESMA VERSAO (evita copia desnecessaria) ---
if (Test-Path $DESTINO) {
    $hashOrigem  = (Get-FileHash $ORIGEM  -Algorithm MD5).Hash
    $hashDestino = (Get-FileHash $DESTINO -Algorithm MD5).Hash
    if ($hashOrigem -eq $hashDestino) {
        Add-Content $LOG "$(Get-Date -Format 'yyyy-MM-dd HH:mm') | OK: ConsoleFlexo ja esta atualizado."
        exit 0
    }
}

# --- COPIA E REGISTRA ---
try {
    $pastaDestino = Split-Path $DESTINO
    if (-not (Test-Path $pastaDestino)) { New-Item -ItemType Directory -Path $pastaDestino | Out-Null }

    Copy-Item -Path $ORIGEM -Destination $DESTINO -Force
    Add-Content $LOG "$(Get-Date -Format 'yyyy-MM-dd HH:mm') | SUCESSO: ConsoleFlexo.gms atualizado de $ORIGEM"
}
catch {
    Add-Content $LOG "$(Get-Date -Format 'yyyy-MM-dd HH:mm') | ERRO ao copiar: $_"
    exit 1
}

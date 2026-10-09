<#
  WuWa Diary - Importador de giros (Convene)            versao 1.0.0

  O QUE ESTE SCRIPT FAZ (e so isso):
    1. Procura o log do jogo (Client.log) no SEU computador e le o arquivo.
    2. Extrai de dentro dele o link do historico de Convene (o jogo grava esse
       link quando voce abre o historico).
    3. Usa esse link para consultar a API OFICIAL da Kuro e baixar os seus giros.
    4. Copia o resultado (JSON) para a sua area de transferencia.

  O QUE ESTE SCRIPT NAO FAZ:
    - NAO envia nada para este site, para o autor, nem para qualquer outro servidor.
      O unico destino de rede e gmserver-api.aki-game2.net (ou .com no servidor CN),
      que e o servidor da propria Kuro. O endereco e validado no codigo.
    - NAO altera nenhum arquivo do jogo (nao mexe em Engine.ini, permissoes, etc).
    - NAO precisa de administrador, NAO escreve no registro, NAO instala nada,
      NAO baixa nem executa outros arquivos.
    - NAO pede senha e NAO le senha. O JSON gerado NAO contem o seu link/token
      (record_id): so UID, servidor e os giros.

  Registro (leitura): so le a lista de programas instalados para achar a pasta
  do Wuthering Waves. Se nao achar, pede o caminho a voce.

  Este arquivo e texto puro: leia tudo antes de rodar. Sao cerca de 300 linhas, sem nada ofuscado.
#>

function Invoke-WuwaDiaryImport {
    $ErrorActionPreference = 'Stop'

    # ---------------- Configuracao (tudo visivel, nada escondido) ----------------
    $ScriptVersion = '1.0.0'
    $PoolTypes     = 1..15      # tipos de banner consultados (veja o resumo no final)
    $DelayMs       = 400        # pausa entre consultas, para nao sobrecarregar a API
    $MaxLogBytes   = 64MB       # le no maximo os ultimos 64 MB do log (o link fica no fim)
    $ApiPath       = '/gacha/record/query'
    $AllowedApi    = @('https://gmserver-api.aki-game2.net', 'https://gmserver-api.aki-game2.com')

    $latin1 = [Text.Encoding]::GetEncoding(28591)   # 1 byte = 1 caractere (sem perdas)

    function Say([string]$Text, [string]$Color = 'Gray') { Write-Host $Text -ForegroundColor $Color }

    # ---------------- Achar os arquivos de log ----------------
    function Get-GameRoots {
        $roots = New-Object System.Collections.Generic.List[string]

        # (a) programas instalados: so entradas com "Wuthering Waves" no nome
        $keys = @(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
        )
        foreach ($k in $keys) {
            Get-ItemProperty -Path $k -ErrorAction SilentlyContinue |
                Where-Object { $_.DisplayName -like '*Wuthering Waves*' } |
                ForEach-Object {
                    foreach ($p in @($_.InstallPath, $_.InstallLocation)) { if ($p) { $roots.Add([string]$p) } }
                }
        }

        # (b) pastas comuns em cada disco fixo
        $rels = @(
            'Wuthering Waves', 'Wuthering Waves Game', 'Games\Wuthering Waves',
            'Program Files\Wuthering Waves', 'Program Files (x86)\Wuthering Waves',
            'Program Files\Epic Games\WutheringWavesj3oFh',
            'Program Files (x86)\Steam\steamapps\common\Wuthering Waves',
            'SteamLibrary\steamapps\common\Wuthering Waves',
            'Steam\steamapps\common\Wuthering Waves'
        )
        foreach ($d in [IO.DriveInfo]::GetDrives()) {
            if ($d.DriveType -eq 'Fixed') {
                foreach ($r in $rels) { $roots.Add((Join-Path $d.Name $r)) }
            }
        }
        return $roots
    }

    function Get-LogFiles([string]$root) {
        $out = @()
        foreach ($base in @($root, [IO.Path]::Combine($root, 'Wuthering Waves Game'))) {
            $candidates = @(
                [IO.Path]::Combine($base, 'Client', 'Saved', 'Logs', 'Client.log'),
                [IO.Path]::Combine($base, 'Client', 'Binaries', 'Win64', 'ThirdParty', 'KrPcSdk_Global', 'KRSDKRes', 'KRSDKWebView', 'debug.log')
            )
            foreach ($f in $candidates) {
                if (Test-Path -LiteralPath $f -PathType Leaf) { $out += Get-Item -LiteralPath $f }
            }
        }
        return $out
    }

    # ---------------- Ler o log e achar o link ----------------
    # Desde a versao 3.4 o jogo grava o Client.log com um XOR por byte:
    #   byte impar -> XOR 0xA5 ; byte par -> XOR 0xEF.
    # Como e uma troca fixa de byte por byte, em vez de decodificar o arquivo todo
    # (lento), codificamos o TRECHO QUE PROCURAMOS, achamos a posicao e decodificamos
    # so uma janelinha de 2 KB ao redor. Logs sem criptografia tambem sao aceitos.
    function Encode-Text([string]$s) {
        $b = New-Object byte[] $s.Length
        for ($i = 0; $i -lt $s.Length; $i++) {
            $p = [int]$s[$i]
            if ($p -band 1) { $b[$i] = $p -bxor 0xEF } else { $b[$i] = $p -bxor 0xA5 }
        }
        return $latin1.GetString($b)
    }

    function Decode-Bytes([byte[]]$b) {
        $o = New-Object byte[] $b.Length
        for ($i = 0; $i -lt $b.Length; $i++) {
            if ($b[$i] -band 1) { $o[$i] = $b[$i] -bxor 0xA5 } else { $o[$i] = $b[$i] -bxor 0xEF }
        }
        return ,$o
    }

    function Find-RecordUrl([string]$path) {
        # abre so para leitura, permitindo que o jogo continue escrevendo no arquivo
        $fs = [IO.File]::Open($path, 'Open', 'Read', ([IO.FileShare]'ReadWrite,Delete'))
        try {
            $len  = $fs.Length
            $take = [int][Math]::Min($len, $MaxLogBytes)
            [void]$fs.Seek($len - $take, 'Begin')
            $buf = New-Object byte[] $take
            $off = 0
            while ($off -lt $take) {
                $n = $fs.Read($buf, $off, $take - $off)
                if ($n -le 0) { break }
                $off += $n
            }
        } finally { $fs.Dispose() }

        $text   = $latin1.GetString($buf, 0, $off)
        $prefix = 'https://aki-gm-resources'
        $regex  = '^https://aki-gm-resources(-oversea)?\.aki-game\.(net|com)/aki/gacha/index\.html#/record\?[^"\s<>\\]+'

        foreach ($mode in @('encoded', 'plain')) {
            $needle = $prefix
            if ($mode -eq 'encoded') { $needle = Encode-Text $prefix }
            $idx = $text.Length - 1
            for ($try = 0; $try -lt 50 -and $idx -ge 0; $try++) {
                $idx = $text.LastIndexOf($needle, $idx, [StringComparison]::Ordinal)
                if ($idx -lt 0) { break }
                $wlen = [Math]::Min(2048, $off - $idx)
                $win  = New-Object byte[] $wlen
                [Array]::Copy($buf, $idx, $win, 0, $wlen)
                if ($mode -eq 'encoded') { $win = Decode-Bytes $win }
                $m = [regex]::Match($latin1.GetString($win), $regex)
                if ($m.Success) { return $m.Value }     # a ocorrencia mais recente vence
                $idx = $idx - 1
            }
        }
        return $null
    }

    function Search-Files($files) {
        foreach ($f in ($files | Sort-Object LastWriteTime -Descending)) {
            Say ('  lendo: ' + $f.FullName)
            $u = Find-RecordUrl $f.FullName
            if ($u) { return [pscustomobject]@{ Url = $u; File = $f } }
        }
        return $null
    }

    function Resolve-UserPath([string]$p) {
        $p = $p.Trim().Trim('"')
        if (-not $p) { return @() }
        if (Test-Path -LiteralPath $p -PathType Leaf) { return @(Get-Item -LiteralPath $p) }
        if (Test-Path -LiteralPath $p -PathType Container) {
            $list = @(Get-LogFiles $p)
            $direct = [IO.Path]::Combine($p, 'Client.log')
            if (Test-Path -LiteralPath $direct -PathType Leaf) { $list += Get-Item -LiteralPath $direct }
            return $list
        }
        return @()
    }

    # ---------------- Consultar a API oficial ----------------
    function Invoke-Pool([string]$apiBase, $q, [string]$lang, [int]$type) {
        $body = [ordered]@{
            cardPoolId   = $q['resources_id']
            cardPoolType = $type
            languageCode = $lang
            playerId     = $q['player_id']
            recordId     = $q['record_id']
            serverId     = $q['svr_id']
        } | ConvertTo-Json -Compress
        $r = Invoke-WebRequest -Uri ($apiBase + $ApiPath) -Method Post -ContentType 'application/json' `
             -Body ([Text.Encoding]::UTF8.GetBytes($body)) -UseBasicParsing -TimeoutSec 30
        if ($r.RawContentStream) { $text = [Text.Encoding]::UTF8.GetString($r.RawContentStream.ToArray()) }
        else { $text = [string]$r.Content }
        return ($text | ConvertFrom-Json)
    }

    # ======================= Execucao =======================
    try {
        Say ''
        Say "=== WuWa Diary - Importador de giros v$ScriptVersion ===" 'Cyan'
        Say 'Este script: le o log do jogo no seu PC, consulta SOMENTE a API oficial da Kuro'
        Say 'e copia seus giros para a area de transferencia. Nao envia nada para outro lugar,'
        Say 'nao altera o jogo e nao precisa de administrador.'
        Say ''
        Say 'Antes de continuar: abra o jogo e entre em Convene > Historico (folheie 2-3 paginas).' 'Yellow'
        Say ''
        [void](Read-Host 'Pressione Enter para continuar (ou feche esta janela para cancelar)')

        # 1) log
        Say ''
        Say '[1/3] Procurando o log do jogo...'
        $found = @{}
        foreach ($root in (Get-GameRoots)) {
            foreach ($f in (Get-LogFiles $root)) { $found[$f.FullName] = $f }
        }
        $hit = $null
        if ($found.Count -gt 0) { $hit = Search-Files @($found.Values) }

        $tries = 0
        while (-not $hit -and $tries -lt 3) {
            $tries++
            if ($found.Count -eq 0) { Say '  Nao achei a pasta do jogo automaticamente.' 'Yellow' }
            else { Say '  Achei o log, mas sem link de historico (abra o historico no jogo e tente de novo).' 'Yellow' }
            $p = Read-Host '  Cole o caminho da pasta do jogo (ou do Client.log). Enter para sair'
            if (-not $p) { break }
            $extra = @(Resolve-UserPath $p)
            if ($extra.Count -eq 0) { Say '  Caminho invalido ou sem log.' 'Yellow'; continue }
            foreach ($f in $extra) { $found[$f.FullName] = $f }
            $hit = Search-Files $extra
        }
        if (-not $hit) {
            throw 'Link do historico nao encontrado. Abra Convene > Historico no jogo, folheie algumas paginas e rode de novo.'
        }
        $ageH = [Math]::Round(((Get-Date) - $hit.File.LastWriteTime).TotalHours, 1)
        Say ('  OK. Log modificado ha ' + $ageH + ' h.') 'Green'
        if ($ageH -gt 6) { Say '  Aviso: o log e antigo; se der erro de link expirado, abra o historico no jogo de novo.' 'Yellow' }

        # 2) validar o link (nada do link e usado sem passar por aqui)
        $url = $hit.Url
        $tld = [regex]::Match($url, '\.aki-game\.(net|com)/').Groups[1].Value
        $apiBase = 'https://gmserver-api.aki-game2.' + $tld
        if ($AllowedApi -notcontains $apiBase) { throw 'Endereco de API fora da lista permitida. Abortado.' }

        $q = @{}
        foreach ($pair in ($url.Substring($url.IndexOf('?') + 1) -split '&')) {
            $kv = $pair -split '=', 2
            if ($kv.Count -eq 2) { $q[$kv[0]] = [Uri]::UnescapeDataString($kv[1]) }
        }
        foreach ($key in @('svr_id', 'player_id', 'record_id', 'resources_id')) {
            if (-not $q.ContainsKey($key) -or $q[$key] -notmatch '^[A-Za-z0-9_-]{1,64}$') {
                throw "Link do historico com formato inesperado (campo $key)."
            }
        }
        $lang = 'en'
        if ($q.ContainsKey('lang') -and $q['lang'] -match '^[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8})?$') { $lang = $q['lang'] }
        Say ('  UID: ' + $q['player_id'] + '   servidor: ' + $tld + '   idioma: ' + $lang)

        # 3) consultar cada banner (uma consulta por tipo, uma de cada vez)
        Say ''
        Say ('[2/3] Consultando ' + $apiBase.Replace('https://', '') + ' (API oficial da Kuro)...')
        $pools = @()
        foreach ($t in $PoolTypes) {
            try {
                $resp = Invoke-Pool $apiBase $q $lang $t
                if ($resp.code -ne 0) {
                    if ($t -eq $PoolTypes[0]) {
                        throw ('A Kuro recusou o link (' + $resp.message + '). Abra o historico no jogo de novo e rode o script outra vez.')
                    }
                    continue
                }
                $recs = @()
                if ($resp.data) { $recs = @($resp.data) }
                if ($recs.Count -gt 0) {
                    $pools += [ordered]@{ type = $t; total = $recs.Count; records = $recs }
                    $five = @($recs | Where-Object { $_.qualityLevel -eq 5 }).Count
                    Say ('  tipo ' + $t + ': ' + $recs.Count + ' giros, ' + $five + ' de 5 estrelas')
                }
            } catch {
                if ($t -eq $PoolTypes[0]) { throw }
                Say ('  tipo ' + $t + ': sem resposta valida (ignorado)') 'DarkGray'
            }
            Start-Sleep -Milliseconds $DelayMs
        }
        if ($pools.Count -eq 0) { throw 'Nenhum giro retornado pela API.' }

        # 4) montar o JSON (sem record_id) e copiar
        $serverName = 'global'
        if ($tld -eq 'com') { $serverName = 'cn' }
        $result = [ordered]@{
            format        = 'wuwa-diary-pulls'
            formatVersion = 1
            script        = $ScriptVersion
            exportedAt    = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
            server        = $serverName
            player        = [ordered]@{ uid = $q['player_id']; serverId = $q['svr_id'] }
            language      = $lang
            scannedTypes  = @($PoolTypes)
            pools         = $pools
        }
        $json = $result | ConvertTo-Json -Depth 8 -Compress

        Say ''
        Say '[3/3] Copiando o resultado...'
        $total = 0
        foreach ($p in $pools) { $total += $p.total }
        try {
            Set-Clipboard -Value $json
            Say ('  Pronto! ' + $total + ' giros copiados para a area de transferencia.') 'Green'
            Say '  Volte ao site e cole (Ctrl+V) na area de importacao.' 'Green'
        } catch {
            $dir = [Environment]::GetFolderPath('Desktop')
            if (-not $dir) { $dir = [Environment]::GetFolderPath('UserProfile') }
            $file = Join-Path $dir ('wuwa-giros-' + $q['player_id'] + '.json')
            [IO.File]::WriteAllText($file, $json, (New-Object Text.UTF8Encoding($false)))
            Say ('  Nao consegui usar a area de transferencia. Salvei em: ' + $file) 'Yellow'
        }
        Say '  Nada foi enviado para fora do seu PC, alem das consultas a API da Kuro.' 'DarkGray'
    }
    catch {
        Say ''
        Say ('[ERRO] ' + $_.Exception.Message) 'Red'
    }
}

Invoke-WuwaDiaryImport

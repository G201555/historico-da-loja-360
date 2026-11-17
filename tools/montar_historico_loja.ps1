# ============================================================
# Historico da Loja - robo de ativacoes
# Consulta o bi_cliente e gera um bloco JS (const ativacoes = {...})
# com o status MAIS RECENTE de Rack Roku TV, Exposicao Especial e
# Testeira para cada loja, e injeta esse bloco dentro do
# historico_da_loja.html (protótipo standalone, sem servidor).
#
# Uso:
#   .\montar_historico_loja.ps1
#
# Formularios usados no bi_cliente:
#   147 = Rack Roku TV        -> situacao: pergunta 990 | foto: 988/1006 | tipo: 981
#   151 = Exposicao Especial  -> situacao: pergunta 1027 | foto: 1025/1031 | tipo: 1019
#   155/156 = Movimentacoes de Testeiras -> acao: pergunta 1113 (INSTALACAO/REMOCAO)
#             foto: 1114 | quantidade: 1115 | parceiro: 929
# Para cada loja pegamos SEMPRE o registro mais recente (maior DATA)
# de cada indicador, para refletir a situacao atual.
# ============================================================
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$htmlFile = Join-Path $root 'historico_da_loja.html'

Write-Host "Historico da Loja - buscando ativacoes (rack / exposicao / testeira)" -ForegroundColor Magenta

$cs = "Server=SEU_SERVIDOR_SQL;Database=bi_cliente;Integrated Security=SSPI;TrustServerCertificate=True;Connection Timeout=20"
$cn = New-Object System.Data.SqlClient.SqlConnection $cs
$cn.Open()

function Invoke-Sql([string]$sql) {
    $cmd = $cn.CreateCommand(); $cmd.CommandTimeout = 300; $cmd.CommandText = $sql
    $da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
    $dt = New-Object System.Data.DataTable
    [void]$da.Fill($dt)
    return , $dt
}
function S([object]$v) { if ($null -eq $v -or $v -is [dbnull]) { return '' }; return "$v".Trim() }
function Foto([object]$v) { $s = S $v; if ($s -like 'http://*') { return 'https://' + $s.Substring(7) }; return $s }

# ---------- Rack Roku TV (form 147) ----------
$rack = Invoke-Sql @"
SELECT f.DATA, f.COD_LOJA, f.COD_PESQUISA,
       a.DES_ALTERNATIVA situacao, tipoAlt.DES_ALTERNATIVA tipo, cod.DES_RESPOSTA codigo, foto.DES_RESPOSTA foto
FROM bi_f_formulario f
JOIN bi_d_formulario_alternativas a ON a.bi_id_formulario_alternativas = f.bi_id_formulario_alternativas
OUTER APPLY (
    SELECT TOP 1 alt.DES_ALTERNATIVA
    FROM bi_f_formulario f2 JOIN bi_d_formulario_alternativas alt ON alt.bi_id_formulario_alternativas = f2.bi_id_formulario_alternativas
    WHERE f2.COD_PESQUISA = f.COD_PESQUISA AND f2.COD_PERGUNTA = 981
) tipoAlt
OUTER APPLY (
    SELECT TOP 1 r.DES_RESPOSTA
    FROM bi_f_formulario f3 JOIN bi_d_formulario_respostas r ON r.bi_id_formulario_respostas = f3.bi_id_formulario_respostas
    WHERE f3.COD_PESQUISA = f.COD_PESQUISA AND f3.COD_PERGUNTA IN (982,997)
) cod
OUTER APPLY (
    SELECT TOP 1 rf.DES_RESPOSTA
    FROM bi_f_formulario f4 JOIN bi_d_formulario_respostas rf ON rf.bi_id_formulario_respostas = f4.bi_id_formulario_respostas
    WHERE f4.COD_PESQUISA = f.COD_PESQUISA AND f4.COD_PERGUNTA IN (988,1006)
) foto
WHERE f.COD_PERGUNTA = 990
ORDER BY f.DATA DESC
"@
$rackFotoFallback = Invoke-Sql @"
SELECT f.COD_LOJA, f.DATA, r.DES_RESPOSTA foto
FROM bi_f_formulario f JOIN bi_d_formulario_respostas r ON r.bi_id_formulario_respostas = f.bi_id_formulario_respostas
WHERE f.COD_PERGUNTA IN (988,1006)
ORDER BY f.DATA DESC
"@

# ---------- Exposicao Especial (form 151) ----------
$exp = Invoke-Sql @"
SELECT f.DATA, f.COD_LOJA, f.COD_PESQUISA,
       a.DES_ALTERNATIVA situacao, tipoAlt.DES_ALTERNATIVA tipo, foto.DES_RESPOSTA foto
FROM bi_f_formulario f
JOIN bi_d_formulario_alternativas a ON a.bi_id_formulario_alternativas = f.bi_id_formulario_alternativas
OUTER APPLY (
    SELECT TOP 1 alt.DES_ALTERNATIVA
    FROM bi_f_formulario f2 JOIN bi_d_formulario_alternativas alt ON alt.bi_id_formulario_alternativas = f2.bi_id_formulario_alternativas
    WHERE f2.COD_PESQUISA = f.COD_PESQUISA AND f2.COD_PERGUNTA = 1019
) tipoAlt
OUTER APPLY (
    SELECT TOP 1 rf.DES_RESPOSTA
    FROM bi_f_formulario f4 JOIN bi_d_formulario_respostas rf ON rf.bi_id_formulario_respostas = f4.bi_id_formulario_respostas
    WHERE f4.COD_PESQUISA = f.COD_PESQUISA AND f4.COD_PERGUNTA IN (1025,1031)
) foto
WHERE f.COD_PERGUNTA = 1027
ORDER BY f.DATA DESC
"@
$expFotoFallback = Invoke-Sql @"
SELECT f.COD_LOJA, f.DATA, r.DES_RESPOSTA foto
FROM bi_f_formulario f JOIN bi_d_formulario_respostas r ON r.bi_id_formulario_respostas = f.bi_id_formulario_respostas
WHERE f.COD_PERGUNTA IN (1025,1031)
ORDER BY f.DATA DESC
"@

# ---------- Testeira (formularios 155/156 - Movimentacoes de Testeiras) ----------
$test = Invoke-Sql @"
SELECT f.DATA, f.COD_LOJA, f.COD_PESQUISA,
       a.DES_ALTERNATIVA acao, parc.DES_ALTERNATIVA parceiro, qtd.DES_RESPOSTA quantidade, foto.DES_RESPOSTA foto
FROM bi_f_formulario f
JOIN bi_d_formulario_alternativas a ON a.bi_id_formulario_alternativas = f.bi_id_formulario_alternativas
OUTER APPLY (
    SELECT TOP 1 palt.DES_ALTERNATIVA
    FROM bi_f_formulario f2 JOIN bi_d_formulario_alternativas palt ON palt.bi_id_formulario_alternativas = f2.bi_id_formulario_alternativas
    WHERE f2.COD_PESQUISA = f.COD_PESQUISA AND f2.COD_PERGUNTA = 929
) parc
OUTER APPLY (
    SELECT TOP 1 r.DES_RESPOSTA
    FROM bi_f_formulario f3 JOIN bi_d_formulario_respostas r ON r.bi_id_formulario_respostas = f3.bi_id_formulario_respostas
    WHERE f3.COD_PESQUISA = f.COD_PESQUISA AND f3.COD_PERGUNTA = 1115
) qtd
OUTER APPLY (
    SELECT TOP 1 rf.DES_RESPOSTA
    FROM bi_f_formulario f4 JOIN bi_d_formulario_respostas rf ON rf.bi_id_formulario_respostas = f4.bi_id_formulario_respostas
    WHERE f4.COD_PESQUISA = f.COD_PESQUISA AND f4.COD_PERGUNTA = 1114
) foto
WHERE f.COD_PERGUNTA = 1113 AND f.COD_FORMULARIO IN (155,156)
ORDER BY f.DATA DESC
"@
$testFotoFallback = Invoke-Sql @"
SELECT f.COD_LOJA, f.DATA, r.DES_RESPOSTA foto
FROM bi_f_formulario f JOIN bi_d_formulario_respostas r ON r.bi_id_formulario_respostas = f.bi_id_formulario_respostas
WHERE f.COD_PERGUNTA = 1114 AND f.COD_FORMULARIO IN (155,156)
ORDER BY f.DATA DESC
"@

$cn.Close()

# ---------- Reducao em memoria: mantem so o registro mais recente por loja ----------
function Get-FotoFallbackMap($dt) {
    $map = @{}
    foreach ($r in $dt.Rows) {
        $k = S $r.COD_LOJA
        if (-not $map.ContainsKey($k)) { $map[$k] = Foto $r.foto }
    }
    return $map
}
$rackFotoMap = Get-FotoFallbackMap $rackFotoFallback
$expFotoMap = Get-FotoFallbackMap $expFotoFallback
$testFotoMap = Get-FotoFallbackMap $testFotoFallback

$ativacoes = [ordered]@{}

foreach ($r in $rack.Rows) {
    $k = S $r.COD_LOJA
    if ($ativacoes.Contains($k) -and $ativacoes[$k].rack) { continue }
    if (-not $ativacoes.Contains($k)) { $ativacoes[$k] = [ordered]@{} }
    $situacao = S $r.situacao
    $foto = Foto $r.foto
    if ([string]::IsNullOrWhiteSpace($foto) -and $rackFotoMap.ContainsKey($k)) { $foto = $rackFotoMap[$k] }
    $ativacoes[$k].rack = [ordered]@{
        situacao = $situacao
        presente = ($situacao -eq 'OK, em perfeito estado.')
        tipo = (S $r.tipo)
        codigo = (S $r.codigo)
        data = (S $r.DATA)
        foto = $foto
    }
}

foreach ($r in $exp.Rows) {
    $k = S $r.COD_LOJA
    if ($ativacoes.Contains($k) -and $ativacoes[$k].exposicao) { continue }
    if (-not $ativacoes.Contains($k)) { $ativacoes[$k] = [ordered]@{} }
    $situacao = S $r.situacao
    $foto = Foto $r.foto
    if ([string]::IsNullOrWhiteSpace($foto) -and $expFotoMap.ContainsKey($k)) { $foto = $expFotoMap[$k] }
    $ativacoes[$k].exposicao = [ordered]@{
        situacao = $situacao
        presente = ($situacao -ne 'Foi removida/descontinuada pela loja.')
        tipo = (S $r.tipo)
        data = (S $r.DATA)
        foto = $foto
    }
}

foreach ($r in $test.Rows) {
    $k = S $r.COD_LOJA
    if ($ativacoes.Contains($k) -and $ativacoes[$k].testeira) { continue }
    if (-not $ativacoes.Contains($k)) { $ativacoes[$k] = [ordered]@{} }
    $acao = S $r.acao
    $foto = Foto $r.foto
    if ([string]::IsNullOrWhiteSpace($foto) -and $testFotoMap.ContainsKey($k)) { $foto = $testFotoMap[$k] }
    $ativacoes[$k].testeira = [ordered]@{
        acao = $acao
        presente = ($acao -like 'INSTALA*')
        parceiro = (S $r.parceiro)
        quantidade = (S $r.quantidade)
        data = (S $r.DATA)
        foto = $foto
    }
}

$lojasComRack = ($ativacoes.Keys | Where-Object { $ativacoes[$_].rack }).Count
$lojasComExp = ($ativacoes.Keys | Where-Object { $ativacoes[$_].exposicao }).Count
$lojasComTest = ($ativacoes.Keys | Where-Object { $ativacoes[$_].testeira }).Count

$json = $ativacoes | ConvertTo-Json -Depth 8 -Compress
$jsBlock = "const ativacoes=$json;"

# ---------- Injeta o bloco no HTML (substitui bloco anterior se ja existir) ----------
$html = [IO.File]::ReadAllText($htmlFile)
$startMarker = '/*__ATIVACOES_START__*/'
$endMarker = '/*__ATIVACOES_END__*/'
$wrapped = "$startMarker`n$jsBlock`n$endMarker"

if ($html.Contains($startMarker)) {
    $pattern = [regex]::Escape($startMarker) + '.*?' + [regex]::Escape($endMarker)
    $html = [regex]::Replace($html, $pattern, [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $wrapped }, [System.Text.RegularExpressions.RegexOptions]::Singleline)
} else {
    $anchor = 'const byStore=new Map();'
    if (-not $html.Contains($anchor)) { throw "Não encontrei o ponto de insercao (const byStore=new Map();) no HTML." }
    $html = $html.Replace($anchor, "$wrapped`n$anchor")
}

[IO.File]::WriteAllText($htmlFile, $html, (New-Object System.Text.UTF8Encoding($false)))

Write-Host "OK -> $htmlFile" -ForegroundColor Green
Write-Host ("  rack: {0} lojas | exposicao: {1} lojas | testeira: {2} lojas" -f $lojasComRack, $lojasComExp, $lojasComTest)

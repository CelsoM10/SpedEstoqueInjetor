using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml;

// ════════════════════════════════════════════════════════════════════════════
//   CONFIGURAÇÃO — altere apenas aqui
// ════════════════════════════════════════════════════════════════════════════
const string PASTA_TRABALHO = @"U:\celso\Projeto SpedEstoqueBlocoK\Trabalho";
const string SUFIXO_SAIDA = "_CORRIGIDO";
const string CODIGO_UNID_PC = "13";   // código da unidade PC no seu sistema
// ════════════════════════════════════════════════════════════════════════════

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

Console.Clear();
Console.WriteLine("╔══════════════════════════════════════════════════════════╗");
Console.WriteLine("║        SPED EFD — Injetor de Estoque (0200 + K200)       ║");
Console.WriteLine("╚══════════════════════════════════════════════════════════╝");
Console.WriteLine();

// ── 1. Verificar / criar pasta de trabalho ─────────────────────────────────
if (!Directory.Exists(PASTA_TRABALHO))
{
    Directory.CreateDirectory(PASTA_TRABALHO);
    Erro($"Pasta criada: {PASTA_TRABALHO}\n" +
         "Coloque a planilha .xlsx e o arquivo SPED .txt nela e execute novamente.");
}
Console.WriteLine($"Pasta de trabalho : {PASTA_TRABALHO}\n");

// ── 2. Localizar planilha Excel ────────────────────────────────────────────
string[] excels = Directory.GetFiles(PASTA_TRABALHO, "*.xlsx");
if (excels.Length == 0)
    Erro("Nenhuma planilha .xlsx encontrada na pasta de trabalho.");

string xlsxPath = excels.Length == 1
    ? excels[0]
    : EscolherArquivo("Escolha a planilha Excel:", excels);

Console.WriteLine($"Planilha          : {Path.GetFileName(xlsxPath)}");

// ── 3. Localizar arquivo SPED ──────────────────────────────────────────────
string[] speds = Directory.GetFiles(PASTA_TRABALHO, "*.txt")
    .Where(f => !Path.GetFileNameWithoutExtension(f).EndsWith(SUFIXO_SAIDA))
    .ToArray();

if (speds.Length == 0)
    Erro("Nenhum arquivo SPED .txt encontrado na pasta de trabalho.");

string spedEntrada = speds.Length == 1
    ? speds[0]
    : EscolherArquivo("Escolha o arquivo SPED:", speds);

string spedSaida = Path.Combine(
    PASTA_TRABALHO,
    Path.GetFileNameWithoutExtension(spedEntrada) + SUFIXO_SAIDA + ".txt");

Console.WriteLine($"SPED entrada      : {Path.GetFileName(spedEntrada)}");
Console.WriteLine($"SPED saida        : {Path.GetFileName(spedSaida)}");
Console.WriteLine();

// ── 4. Processar ───────────────────────────────────────────────────────────
Console.WriteLine("Lendo planilha...");
var itens = LerItens(xlsxPath);
var estoques = LerEstoques(xlsxPath);
Console.WriteLine($"  {itens.Count} produto(s) unico(s) encontrado(s)");
Console.WriteLine($"  {estoques.Count} linha(s) de estoque encontrada(s)");

Console.WriteLine("Lendo SPED e limpando assinaturas antigas...");
var linhasSped = new List<string>();

// Suporte para a codificação padrão do SPED (ISO-8859-1) no .NET
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
var encodingSped = Encoding.GetEncoding("ISO-8859-1");

using (var reader = new StreamReader(spedEntrada, encodingSped))
{
    string? linha;
    while ((linha = reader.ReadLine()) != null)
    {
        // Altera a finalidade do arquivo no registro 0000 para 1 (Retificador)
        if (linha.StartsWith("|0000|"))
        {
            var campos = linha.Split('|');
            if (campos.Length > 3)
            {
                campos[3] = "1"; // Força Retificação
            }
            linha = string.Join("|", campos);
        }

        linhasSped.Add(linha);

        // Se encontrou o fim dos registros do SPED, interrompe para descartar o lixo da assinatura digital antiga
        if (linha.StartsWith("|9999|"))
        {
            break;
        }
    }
}
Console.WriteLine($"  {linhasSped.Count} linha(s) válidas carregadas.");

Console.WriteLine("Injetando 0200 e K200...");
var resultado = InjetarRegistros(linhasSped, itens, estoques, CODIGO_UNID_PC);

Console.WriteLine("Recalculando contadores...");
resultado = RecalcularContadores(resultado);

// Salva o arquivo final respeitando rigorosamente a codificação exigida pelo PVA
File.WriteAllLines(spedSaida, resultado, encodingSped);

Console.WriteLine();
Console.WriteLine("╔══════════════════════════════════════════════════════════╗");
Console.WriteLine("║                    CONCLUIDO COM SUCESSO!                ║");
Console.WriteLine("╚══════════════════════════════════════════════════════════╝");
Console.WriteLine($"\nArquivo gerado: {spedSaida}");
Console.WriteLine($"Total de linhas: {resultado.Count}");
Console.WriteLine("\nPressione qualquer tecla para fechar...");
Console.ReadKey();

// ════════════════════════════════════════════════════════════════════════════
//  LEITURA DO EXCEL
// ════════════════════════════════════════════════════════════════════════════

static List<(string Cod, string Descr, string Unid, string Tipo, string Ncm)>
    LerItens(string path)
{
    var vistos = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    var lista = new List<(string, string, string, string, string)>();
    using var pkg = new ExcelPackage(new FileInfo(path));
    var ws = pkg.Workbook.Worksheets[0];
    int rows = ws.Dimension?.Rows ?? 1;
    for (int r = 2; r <= rows; r++)
    {
        string cod = Cel(ws, r, 1);
        if (string.IsNullOrWhiteSpace(cod) || vistos.Contains(cod)) continue;
        vistos.Add(cod);
        lista.Add((
            cod,
            Cel(ws, r, 2),
            Cel(ws, r, 3),
            ExtrairTipo(Cel(ws, r, 4)),
            NcmLimpo(Cel(ws, r, 5))
        ));
    }
    return lista;
}

static List<(string Cod, decimal Qtd, int IndEst, string Cnpj)>
    LerEstoques(string path)
{
    var lista = new List<(string, decimal, int, string)>();
    using var pkg = new ExcelPackage(new FileInfo(path));
    var ws = pkg.Workbook.Worksheets[0];
    int rows = ws.Dimension?.Rows ?? 1;
    for (int r = 2; r <= rows; r++)
    {
        string cod = Cel(ws, r, 1);
        if (string.IsNullOrWhiteSpace(cod)) continue;

        string qtdTexto = Cel(ws, r, 6).Replace(".", "").Replace(",", ".");
        decimal.TryParse(qtdTexto, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal qtd);

        string tipo = Cel(ws, r, 7);
        int ind = tipo.Contains("terceiro", StringComparison.OrdinalIgnoreCase) ? 1 : 0;
        string cnpj = ind == 1 ? Cel(ws, r, 8) : "";
        lista.Add((cod, qtd, ind, cnpj));
    }
    return lista;
}

// ════════════════════════════════════════════════════════════════════════════
//  INJEÇÃO NO SPED
// ════════════════════════════════════════════════════════════════════════════

static List<string> InjetarRegistros(
    List<string> linhas,
    List<(string Cod, string Descr, string Unid, string Tipo, string Ncm)> itens,
    List<(string Cod, decimal Qtd, int IndEst, string Cnpj)> estoques,
    string codUnidPc)
{
    string dataFinal = DataFinalPeriodo(linhas);

    // Unidades já cadastradas no 0190
    var unidsCadastradas = linhas
        .Where(l => l.StartsWith("|0190|"))
        .Select(l => l.Split('|'))
        .Where(p => p.Length > 2)
        .Select(p => p[2])
        .ToHashSet(StringComparer.OrdinalIgnoreCase);

    // Mapa CNPJ limpo → código participante (0150)
    var mapaPart = linhas
        .Where(l => l.StartsWith("|0150|"))
        .Select(l => l.Split('|'))
        .Where(p => p.Length > 5 && !string.IsNullOrWhiteSpace(p[5]))
        .GroupBy(p => CnpjLimpo(p[5]))
        .ToDictionary(g => g.Key, g => g.First()[2],
            StringComparer.OrdinalIgnoreCase);

    // Novas unidades de medida para o 0190
    var novas0190 = itens
        .Select(i => i.Unid)
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .Where(u => !string.IsNullOrWhiteSpace(u) && !unidsCadastradas.Contains(u))
        .Select(u => $"|0190|{codUnidPc}|{u}|")
        .ToList();

    // Registros 0200
    // Registros 0200 corrigidos para conter exatamente 13 campos (14 pipes no total)
    var novos0200 = itens
        .Select(i => $"|0200|{i.Cod}|{i.Descr}|||{i.Unid}|{i.Tipo}|{i.Ncm}||||||")
        .ToList();

    // Registros K200
    var novosK200 = estoques.Select(e =>
    {
        string qtdFmt = e.Qtd.ToString("F3", new CultureInfo("pt-BR"));
        string codPart = "";
        if (e.IndEst == 1 && !string.IsNullOrWhiteSpace(e.Cnpj))
        {
            string cnpjL = CnpjLimpo(e.Cnpj);
            codPart = mapaPart.TryGetValue(cnpjL, out string? c) ? c : cnpjL;
        }
        return $"|K200|{dataFinal}|{e.Cod}|{qtdFmt}|{e.IndEst}|{codPart}|";
    }).ToList();

    // ── Montar resultado ───────────────────────────────────────────────────
    var res = new List<string>(linhas.Count + novos0200.Count + novosK200.Count + 20);
    bool inseriu0190 = novas0190.Count == 0;
    bool inseriu0200 = false;
    bool inseriuK200 = false;

    foreach (var linha in linhas)
    {
        string reg = Reg(linha);

        // Adicionar novas unidades 0190 antes do primeiro 0200
        if (!inseriu0190 && reg == "0200")
        {
            foreach (var n in novas0190) res.Add(n);
            inseriu0190 = true;
        }

        // Pular 0200 antigos
        if (reg == "0200") continue;

        // Inserir novos 0200 quando bloco 0200 terminar (chega no 0400 ou 0990)
        if (!inseriu0200 && (reg == "0400" || reg == "0990"))
        {
            foreach (var n in novos0200) res.Add(n);
            inseriu0200 = true;
        }

        // Pular K200 antigos
        if (reg == "K200") continue;

        // Se encontrar abertura de bloco K vazia (K001 com '1'), altera para '0' (com dados)
        if (reg == "K001")
        {
            var camposK = linha.Split('|');
            if (camposK.Length > 2 && camposK[2] == "1")
            {
                camposK[2] = "0";
                res.Add(string.Join("|", camposK));
                continue;
            }
        }

        // Inserir novos K200 antes do K990
        if (!inseriuK200 && reg == "K990")
        {
            foreach (var n in novosK200) res.Add(n);
            inseriuK200 = true;
        }

        res.Add(linha);
    }

    // Segurança: pontos de inserção não encontrados
    if (!inseriu0200) InserirAntes(res, "0990", novos0200);
    if (!inseriuK200) InserirAntes(res, "K990", novosK200);

    Console.WriteLine($"  {novos0200.Count} registro(s) 0200 injetado(s)");
    Console.WriteLine($"  {novosK200.Count} registro(s) K200 injetado(s)");
    if (novas0190.Count > 0)
        Console.WriteLine($"  {novas0190.Count} unidade(s) nova(s) adicionada(s) ao 0190");

    return res;
}

// ════════════════════════════════════════════════════════════════════════════
//  RECÁLCULO DE CONTADORES
// ════════════════════════════════════════════════════════════════════════════

static List<string> RecalcularContadores(List<string> linhas)
{
    (string, string)[] blocos =
    [
        ("0000","0990"), ("B001","B990"), ("C001","C990"), ("D001","D990"),
        ("E001","E990"), ("G001","G990"), ("H001","H990"), ("K001","K990"),
        ("1001","1990")
    ];
    foreach (var (ini, fim) in blocos)
        CorrigirBloco(linhas, ini, fim);

    return ReconstruirBloco9(linhas);
}

static void CorrigirBloco(List<string> linhas, string ini, string fim)
{
    int idxIni = -1, idxFim = -1;
    for (int i = 0; i < linhas.Count; i++)
    {
        string r = Reg(linhas[i]);
        if (r == ini && idxIni < 0) idxIni = i;
        if (r == fim) { idxFim = i; break; }
    }
    if (idxIni < 0 || idxFim < 0) return;
    var p = linhas[idxFim].TrimEnd('\r', '\n').Split('|');
    if (p.Length > 2) p[2] = (idxFim - idxIni + 1).ToString();
    linhas[idxFim] = string.Join("|", p);
}

static List<string> ReconstruirBloco9(List<string> linhas)
{
    string[] regs9 = ["9001", "9900", "9990", "9999"];
    var sem9 = linhas.Where(l => !regs9.Contains(Reg(l))).ToList();

    var contagens = sem9
        .GroupBy(l => Reg(l))
        .ToDictionary(g => g.Key, g => g.Count());

    // Inicializa os registros obrigatórios do próprio bloco 9 para entrarem no 9900
    contagens["9001"] = 1;
    contagens["9990"] = 1;
    contagens["9999"] = 1;

    int totalRegistrosDiferentes = contagens.Count + 1; // +1 para incluir o próprio 9900 na contagem
    contagens["9900"] = totalRegistrosDiferentes;

    var linhas9900 = contagens
        .OrderBy(kv => kv.Key)
        .Select(kv => $"|9900|{kv.Key}|{kv.Value}|")
        .ToList();

    var bloco9 = new List<string> { "|9001|0|" };
    bloco9.AddRange(linhas9900);

    // Regra do PVA: O 9990 conta TODAS as linhas do bloco 9, INCLUINDO a própria linha do 9990.
    // Como a lista 'bloco9' atual tem (9001 + as linhas 9900), somamos +1 que representa a linha do 9990 que estamos inserindo agora.
    int qtd9990 = bloco9.Count + 2;
    bloco9.Add($"|9990|{qtd9990}|");

    // Total geral do arquivo (linhas dos outros blocos + linhas do bloco 9 + a própria linha do 9999)
    int qtd9999 = sem9.Count + bloco9.Count + 1;
    bloco9.Add($"|9999|{qtd9999}|");

    sem9.AddRange(bloco9);
    return sem9;
}

// ════════════════════════════════════════════════════════════════════════════
//  AUXILIARES
// ════════════════════════════════════════════════════════════════════════════

static string Reg(string linha)
{
    var p = linha.TrimEnd('\r', '\n').Split('|');
    return p.Length > 1 ? p[1] : "";
}

static string Cel(ExcelWorksheet ws, int row, int col)
    => ws.Cells[row, col].Text?.Trim() ?? "";

static string ExtrairTipo(string valor)
{
    var m = Regex.Match(valor, @"^\d+");
    return m.Success ? m.Value.PadLeft(2, '0') : "00";
}

static string NcmLimpo(string ncm)
    => Regex.Replace(ncm ?? "", @"[.\-/\s]", "");

static string CnpjLimpo(string cnpj)
    => Regex.Replace(cnpj ?? "", @"[.\-/\s]", "");

static string DataFinalPeriodo(List<string> linhas)
{
    foreach (var l in linhas)
        if (l.StartsWith("|0000|"))
        {
            var p = l.Split('|');
            if (p.Length > 5) return p[5];
        }
    return DateTime.Now.ToString("ddMMyyyy");
}

static void InserirAntes(List<string> linhas, string reg, List<string> novas)
{
    int idx = linhas.FindIndex(l => Reg(l) == reg);
    if (idx >= 0) linhas.InsertRange(idx, novas);
    else linhas.AddRange(novas);
}

static string EscolherArquivo(string titulo, string[] arquivos)
{
    Console.WriteLine(titulo);
    for (int i = 0; i < arquivos.Length; i++)
        Console.WriteLine($"  [{i + 1}] {Path.GetFileName(arquivos[i])}");
    Console.Write("Digite o numero: ");
    int escolha = int.Parse(Console.ReadLine() ?? "1") - 1;
    return arquivos[Math.Clamp(escolha, 0, arquivos.Length - 1)];
}

static void Erro(string msg)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine($"\nERRO: {msg}");
    Console.ResetColor();
    Console.WriteLine("\nPressione qualquer tecla para fechar...");
    Console.ReadKey();
    Environment.Exit(1);
}

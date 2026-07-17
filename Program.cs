using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace SpedEstoqueInjetor
{
    internal class Program
    {
        // ════════════════════════════════════════════════════════════════════════════
        //   CONFIGURAÇÃO — altere apenas aqui
        // ════════════════════════════════════════════════════════════════════════════
        const string PASTA_TRABALHO = @"U:\celso\Projeto SpedEstoqueBlocoK\Trabalho";
        const string SUFIXO_SAIDA = "_CORRIGIDO";
        // ════════════════════════════════════════════════════════════════════════════

        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            Console.Clear();
            Console.WriteLine("╔══════════════════════════════════════════════════════════╗");
            Console.WriteLine("║          SPED EFD — Injetor Cirúrgico do Bloco K         ║");
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

            // ── 2. Localizar planilha Excel (Automático) ───────────────────────────────
            string[] excels = Directory.GetFiles(PASTA_TRABALHO, "*.xlsx");
            if (excels.Length == 0)
                Erro("Nenhuma planilha .xlsx encontrada na pasta de trabalho.");

            string xlsxPath = excels[0];
            Console.WriteLine($"Planilha          : {Path.GetFileName(xlsxPath)}");

            // ── 3. Localizar arquivo SPED (Automático) ─────────────────────────────────
            string[] speds = Directory.GetFiles(PASTA_TRABALHO, "*.txt")
                .Where(f => !Path.GetFileNameWithoutExtension(f).EndsWith(SUFIXO_SAIDA))
                .ToArray();

            if (speds.Length == 0)
                Erro("Nenhum arquivo SPED .txt encontrado na pasta de trabalho.");

            string spedEntrada = speds[0];
            string spedSaida = Path.Combine(
                PASTA_TRABALHO,
                Path.GetFileNameWithoutExtension(spedEntrada) + SUFIXO_SAIDA + ".txt");
            string relatorioPendencias = Path.Combine(
                PASTA_TRABALHO,
                Path.GetFileNameWithoutExtension(spedEntrada) + SUFIXO_SAIDA + "_PENDENCIAS.txt");

            Console.WriteLine($"SPED entrada      : {Path.GetFileName(spedEntrada)}");
            Console.WriteLine($"SPED saida        : {Path.GetFileName(spedSaida)}");
            Console.WriteLine();

            // ── 4. Ler Linhas Originais do SPED ────────────────────────────────────────
            Console.WriteLine("Lendo SPED original...");
            var linhasOriginais = new List<string>();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var encodingSped = Encoding.GetEncoding("ISO-8859-1");

            using (var reader = new StreamReader(spedEntrada, encodingSped))
            {
                string? linha;
                while ((linha = reader.ReadLine()) != null)
                {
                    linha = linha.Trim();
                    if (string.IsNullOrWhiteSpace(linha)) continue;

                    linhasOriginais.Add(linha);
                    if (linha.StartsWith("|9999|")) break;
                }
            }

            // ── 5. Mapear o Cadastro Oficial de Produtos (Registro 0200) Existente ──────
            // Chave: Código do Item formatado | Valor: Linha Completa do 0200
            // O 0200 já veio validado e transmitido — NUNCA é recriado ou alterado aqui,
            // serve apenas como fonte de verdade para o cruzamento com a planilha.
            var produtosDoTxt = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var linha in linhasOriginais)
            {
                if (Reg(linha) == "0200")
                {
                    var campos = linha.Split('|');
                    if (campos.Length > 2)
                    {
                        string codItem = campos[2].Trim();
                        if (!produtosDoTxt.ContainsKey(codItem))
                        {
                            produtosDoTxt.Add(codItem, linha);
                        }
                    }
                }
            }
            Console.WriteLine($"  {produtosDoTxt.Count} produto(s) mapeado(s) diretamente do arquivo original.");

            // Mapear também os participantes (0150) para tratamento de estoque em terceiros
            var mapaPart = linhasOriginais
                .Where(l => Reg(l) == "0150")
                .Select(l => l.Split('|'))
                .Where(p => p.Length > 5 && !string.IsNullOrWhiteSpace(p[5]))
                .GroupBy(p => CnpjLimpo(p[5]))
                .ToDictionary(g => g.Key, g => g.First()[2], StringComparer.OrdinalIgnoreCase);

            // Obtém as datas de Início e Fim da apuração do registro 0000 original
            var (dataInicio, dataFim) = ObterDatasPeriodo(linhasOriginais);

            // ── 6. Ler Planilha Excel e Validar Cruzamento de Dados ────────────────────
            Console.WriteLine("Processando planilha de inventário...");
            var estoquesPlanilha = LerEstoques(xlsxPath);
            var novosK200 = new List<string>();

            // Pendências: nunca injetamos dado que pode invalidar o arquivo no PVA.
            // Tudo que não casa 100% vira linha de relatório para revisão manual.
            var itensSemCadastro0200 = new List<(string Cod, decimal Qtd)>();
            var itensSemParticipante0150 = new List<(string Cod, decimal Qtd, string Cnpj)>();

            foreach (var est in estoquesPlanilha)
            {
                // Validação Cruzada: O produto do inventário existe no 0200 original?
                // Repare que isso NÃO exige movimento no mês — só cadastro — então
                // itens parados em estoque mas sem giro no período entram normalmente.
                if (!produtosDoTxt.ContainsKey(est.Cod.Trim()))
                {
                    itensSemCadastro0200.Add((est.Cod, est.Qtd));
                    continue;
                }

                string qtdFmt = est.Qtd.ToString("F3", new CultureInfo("pt-BR"));
                string codPart = "";

                if (est.IndEst == 1 && !string.IsNullOrWhiteSpace(est.Cnpj))
                {
                    string cnpjL = CnpjLimpo(est.Cnpj);
                    if (mapaPart.TryGetValue(cnpjL, out string? c))
                    {
                        codPart = c;
                    }
                    else
                    {
                        // Participante de terceiro não está no 0150 do arquivo original.
                        // Gravar CNPJ cru no COD_PART geraria referência inválida no PVA,
                        // então o item fica de fora do K200 e vai para o relatório.
                        itensSemParticipante0150.Add((est.Cod, est.Qtd, est.Cnpj));
                        continue;
                    }
                }

                // Monta o K200 referenciado à data limite da apuração (dataFim)
                novosK200.Add($"|K200|{dataFim}|{est.Cod.Trim()}|{qtdFmt}|{est.IndEst}|{codPart}|");
            }

            if (itensSemCadastro0200.Count > 0)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"  [Aviso] {itensSemCadastro0200.Count} item(ns) da planilha ignorados por não constarem no 0200 original.");
                Console.ResetColor();
            }
            if (itensSemParticipante0150.Count > 0)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"  [Aviso] {itensSemParticipante0150.Count} item(ns) de estoque em terceiros ignorados por participante ausente no 0150.");
                Console.ResetColor();
            }
            Console.WriteLine($"  {novosK200.Count} linha(s) de K200 preparadas para injeção.");

            GravarRelatorioPendencias(relatorioPendencias, itensSemCadastro0200, itensSemParticipante0150);

            // ── 7. Limpeza Cirúrgica do Bloco K antigo e Reordenamento Estrutural ──────
            Console.WriteLine("Preparando estrutura limpa para o merge...");
            var linhasSpedLimpo = new List<string>();
            int linhasKRemovidas = 0;

            foreach (var linha in linhasOriginais)
            {
                string reg = Reg(linha);

                // Força o 0000 como retificador (arquivo já foi transmitido antes)
                if (reg == "0000")
                {
                    var campos = linha.Split('|');
                    if (campos.Length > 3) campos[3] = "1"; // 1 = Retificadora
                    linhasSpedLimpo.Add(string.Join("|", campos));
                    continue;
                }

                // Deleta QUALQUER resquício de Bloco K, não só os tipos previstos (K001,
                // K100, K200, K280, K990). Arquivos de origem (ex: sistema da bomba de
                // combustível) podem trazer registros K não-padrão ou incompletos — como
                // "K010", que não existe no leiaute oficial — e uma lista branca fechada
                // deixaria isso passar, quebrando a hierarquia quando o Bloco K novo é
                // injetado. Este programa é o único responsável por montar o Bloco K
                // inteiro do zero, então a limpeza precisa ser total.
                if (reg.StartsWith("K"))
                {
                    linhasKRemovidas++;
                    continue;
                }

                linhasSpedLimpo.Add(linha);
            }
            if (linhasKRemovidas > 0)
                Console.WriteLine($"  {linhasKRemovidas} linha(s) residual(is) de Bloco K removida(s) do arquivo original.");

            // ── 8. Executar Injeção Cirúrgica através do Engine ────────────────────────
            Console.WriteLine("Injetando Bloco K...");

            var engine = new SpedMergeEngine(linhasSpedLimpo, novosK200, dataInicio, dataFim);
            var resultado = engine.Processar();

            // ── 9. Recalcular Contadores ────────────────────────────────────────────────
            Console.WriteLine("Recalculando contadores internos e Bloco 9...");
            resultado = RecalcularContadores(resultado);

            // ── 9.5 Trava de Segurança Fiscal: TODO item do Bloco K precisa existir no 0200 ──
            // Isso não é uma checagem redundante do passo 6: aqui varremos o arquivo FINAL
            // já montado, então pega qualquer inconsistência introduzida em qualquer etapa
            // posterior (merge, reordenação, etc.), não só o que veio direto da planilha.
            Console.WriteLine("Validando integridade referencial do Bloco K contra o 0200...");
            var codItensValidos = new HashSet<string>(produtosDoTxt.Keys, StringComparer.OrdinalIgnoreCase);
            var errosValidacao = ValidarIntegridadeBlocoK(resultado, codItensValidos);

            if (errosValidacao.Count > 0)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine();
                Console.WriteLine("╔══════════════════════════════════════════════════════════╗");
                Console.WriteLine("║   VALIDAÇÃO FALHOU — GERAÇÃO DO SPED FOI ABORTADA         ║");
                Console.WriteLine("╚══════════════════════════════════════════════════════════╝");
                Console.WriteLine("Os seguintes registros do Bloco K referenciam COD_ITEM ausente no 0200:");
                foreach (var e in errosValidacao)
                    Console.WriteLine($"  {e}");
                Console.ResetColor();
                Console.WriteLine("\nNenhum arquivo foi gravado. Corrija a origem e rode novamente.");
                Console.WriteLine("\nPressione qualquer tecla para fechar...");
                Console.ReadKey();
                Environment.Exit(1);
            }
            Console.WriteLine("  OK — todos os COD_ITEM do Bloco K estão previamente cadastrados no 0200.");

            using (var writer = new StreamWriter(spedSaida, false, encodingSped))
            {
                writer.NewLine = "\r\n";
                foreach (var l in resultado)
                {
                    if (string.IsNullOrWhiteSpace(l)) continue;
                    writer.WriteLine(l.Trim());
                }
            }

            Console.WriteLine();
            Console.WriteLine("╔══════════════════════════════════════════════════════════╗");
            Console.WriteLine("║            CONCLUIDO CIRURGICAMENTE COM SUCESSO!         ║");
            Console.WriteLine("╚══════════════════════════════════════════════════════════╝");
            Console.WriteLine($"\nArquivo gerado: {spedSaida}");
            Console.WriteLine($"Total de linhas: {resultado.Count}");
            if (itensSemCadastro0200.Count > 0 || itensSemParticipante0150.Count > 0)
                Console.WriteLine($"Relatório de pendências: {relatorioPendencias}");
            Console.WriteLine("\nPressione qualquer tecla para fechar...");
            Console.ReadKey();
        }

        // ════════════════════════════════════════════════════════════════════════════
        //  AUXILIARES DE LEITURA
        // ════════════════════════════════════════════════════════════════════════════

        static List<(string Cod, decimal Qtd, int IndEst, string Cnpj)> LerEstoques(string path)
        {
            var acumulado = new Dictionary<string, (string Cod, decimal Qtd, int IndEst, string Cnpj)>(
                StringComparer.OrdinalIgnoreCase);

            using var pkg = new ExcelPackage(new FileInfo(path));
            var ws = pkg.Workbook.Worksheets[0];
            int rows = ws.Dimension?.Rows ?? 1;

            for (int r = 2; r <= rows; r++)
            {
                string cod = ws.Cells[r, 1].Text?.Trim() ?? "";
                if (string.IsNullOrWhiteSpace(cod)) continue;

                string qtdTexto = ws.Cells[r, 6].Text?.Trim().Replace(".", "").Replace(",", ".") ?? "0";
                decimal.TryParse(qtdTexto, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal qtd);

                string tipo = ws.Cells[r, 7].Text?.Trim() ?? "";
                int ind = tipo.Contains("terceiro", StringComparison.OrdinalIgnoreCase) ? 1 : 0;
                string cnpj = ind == 1 ? ws.Cells[r, 8].Text?.Trim() ?? "" : "";

                string chave = $"{cod.Trim()}|{ind}|{Regex.Replace(cnpj, @"[.\-/\s]", "")}";

                if (acumulado.TryGetValue(chave, out var existente))
                    acumulado[chave] = existente with { Qtd = existente.Qtd + qtd };
                else
                    acumulado[chave] = (cod, qtd, ind, cnpj);
            }

            return acumulado.Values.ToList();
        }

        static void GravarRelatorioPendencias(
            string path,
            List<(string Cod, decimal Qtd)> semCadastro,
            List<(string Cod, decimal Qtd, string Cnpj)> semParticipante)
        {
            if (semCadastro.Count == 0 && semParticipante.Count == 0)
            {
                if (File.Exists(path)) File.Delete(path);
                return;
            }

            var sb = new StringBuilder();
            sb.AppendLine("RELATÓRIO DE PENDÊNCIAS — ITENS NÃO INJETADOS NO BLOCO K");
            sb.AppendLine($"Gerado em: {DateTime.Now:dd/MM/yyyy HH:mm:ss}");
            sb.AppendLine(new string('-', 70));

            if (semCadastro.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("1) ITENS SEM CADASTRO NO 0200 DO ARQUIVO ORIGINAL");
                sb.AppendLine("   (código da planilha não encontrado no SPED — provável divergência de código)");
                foreach (var it in semCadastro.OrderBy(i => i.Cod))
                    sb.AppendLine($"   COD_ITEM={it.Cod}  QTD={it.Qtd:F3}");
            }

            if (semParticipante.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("2) ESTOQUE EM PODER DE TERCEIROS SEM PARTICIPANTE NO 0150");
                sb.AppendLine("   (CNPJ da planilha não encontrado no cadastro de participantes do SPED)");
                foreach (var it in semParticipante.OrderBy(i => i.Cod))
                    sb.AppendLine($"   COD_ITEM={it.Cod}  QTD={it.Qtd:F3}  CNPJ={it.Cnpj}");
            }

            File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
        }

        // ════════════════════════════════════════════════════════════════════════════
        //  TRAVA DE SEGURANÇA FISCAL — Registro 0200 é registro-pai de todo Bloco K.
        //  A legislação (Guia Prático EFD-ICMS/IPI) não permite nenhum COD_ITEM no
        //  Bloco K que não esteja previamente cadastrado no 0200. Esta validação varre
        //  o arquivo final já montado e barra a gravação se achar qualquer violação.
        // ════════════════════════════════════════════════════════════════════════════

        // Mapa: registro do Bloco K -> índices (no array Split('|')) onde há um COD_ITEM.
        // Hoje só K200 é gerado por este programa. Se no futuro forem adicionados
        // registros de produção (K220, K230, K250...), basta declarar aqui os índices
        // dos campos correspondentes que esta validação passa a cobrir também.
        static readonly Dictionary<string, int[]> CamposCodItemPorRegistroK = new()
        {
            { "K200", new[] { 3 } }, // |K200|DT_EST|COD_ITEM|QTD|IND_EST|COD_PART|
        };

        static List<string> ValidarIntegridadeBlocoK(List<string> linhas, HashSet<string> codItensValidos)
        {
            var erros = new List<string>();

            foreach (var linha in linhas)
            {
                string reg = Reg(linha);
                if (!reg.StartsWith("K")) continue;
                if (!CamposCodItemPorRegistroK.TryGetValue(reg, out int[]? indices)) continue;

                var campos = linha.Split('|');
                foreach (int idx in indices)
                {
                    if (idx >= campos.Length) continue;
                    string codItem = campos[idx].Trim();
                    if (string.IsNullOrWhiteSpace(codItem)) continue;

                    if (!codItensValidos.Contains(codItem))
                        erros.Add($"{reg}: COD_ITEM='{codItem}' não encontrado no 0200  →  linha: {linha}");
                }
            }

            return erros;
        }

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

            var camposFim = linhas[idxFim].Split('|');
            if (camposFim.Length > 2) camposFim[2] = (idxFim - idxIni + 1).ToString();
            linhas[idxFim] = string.Join("|", camposFim);
        }

        static List<string> ReconstruirBloco9(List<string> linhas)
        {
            string[] regs9 = ["9001", "9900", "9990", "9999"];
            var sem9 = linhas.Where(l => !regs9.Contains(Reg(l))).ToList();

            var contagens = sem9
                .GroupBy(l => Reg(l))
                .ToDictionary(g => g.Key, g => g.Count());

            contagens["9001"] = 1;
            contagens["9990"] = 1;
            contagens["9999"] = 1;
            int totalRegistrosDiferentes = contagens.Count + 1;
            contagens["9900"] = totalRegistrosDiferentes;

            var linhas9900 = contagens
                .OrderBy(kv => kv.Key)
                .Select(kv => $"|9900|{kv.Key}|{kv.Value}|")
                .ToList();

            var bloco9 = new List<string> { "|9001|0|" };
            bloco9.AddRange(linhas9900);

            int qtd9990 = bloco9.Count + 2;
            bloco9.Add($"|9990|{qtd9990}|");

            int qtd9999 = sem9.Count + bloco9.Count + 1;
            bloco9.Add($"|9999|{qtd9999}|");

            sem9.AddRange(bloco9);
            return sem9;
        }

        static string Reg(string linha)
        {
            var p = linha.Split('|');
            return p.Length > 1 ? p[1] : "";
        }

        static string CnpjLimpo(string cnpj) => Regex.Replace(cnpj ?? "", @"[.\-/\s]", "");

        static (string Inicio, string Fim) ObterDatasPeriodo(List<string> linhas)
        {
            foreach (var l in linhas)
                if (l.StartsWith("|0000|"))
                {
                    var p = l.Split('|');
                    if (p.Length > 5) return (p[4], p[5]); // p[4] = DT_INI, p[5] = DT_FIN
                }
            return ("01102025", "31102025"); // Valores padrão de contingência
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
    }

    internal class LyraLatin1Encoding : System.Text.Encoding
    {
        private readonly Encoding _iso = Encoding.GetEncoding("ISO-8859-1");
        public override byte[] GetPreamble() => Array.Empty<byte>();
        public override int GetByteCount(char[] chars, int index, int count) => _iso.GetByteCount(chars, index, count);
        public override int GetBytes(char[] chars, int charIndex, int charCount, byte[] bytes, int byteIndex) => _iso.GetBytes(chars, charIndex, charCount, bytes, byteIndex);
        public override int GetCharCount(byte[] bytes, int index, int count) => _iso.GetCharCount(bytes, index, count);
        public override int GetChars(byte[] bytes, int byteIndex, int byteCount, char[] chars, int charIndex) => _iso.GetChars(bytes, byteIndex, byteCount, chars, charIndex);
        public override int GetMaxByteCount(int charCount) => _iso.GetMaxByteCount(charCount);
        public override int GetMaxCharCount(int byteCount) => _iso.GetMaxCharCount(byteCount);
    }
}

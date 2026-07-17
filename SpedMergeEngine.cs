using System;
using System.Collections.Generic;
using System.Linq;

namespace SpedEstoqueInjetor
{
    public class SpedMergeEngine
    {
        private readonly List<SpedRegistro> _registrosOriginais;
        private readonly List<string> _novosK200;
        private readonly string _dataInicio;
        private readonly string _dataFim;

        public SpedMergeEngine(List<string> linhasSped, List<string> novosK200, string dataInicio, string dataFim)
        {
            _registrosOriginais = linhasSped.Select(l => new SpedRegistro(l)).ToList();
            _novosK200 = novosK200;
            _dataInicio = dataInicio;
            _dataFim = dataFim;
        }

        public List<string> Processar()
        {
            var resultadoCru = new List<string>();

            // 1. GARANTIA ABSOLUTA DO TOPO (0000)
            var reg0000 = _registrosOriginais.FirstOrDefault(r => r.Registro == "0000");
            if (reg0000 != null) resultadoCru.Add(reg0000.ConteudoCompleto);

            // 2. PROCESSA O RESTANTE DOS BLOCOS ORIGINAIS
            for (int i = 0; i < _registrosOriginais.Count; i++)
            {
                var regAtual = _registrosOriginais[i];

                if (regAtual.Registro == "0000" || regAtual.Registro == "K200")
                    continue;

                resultadoCru.Add(regAtual.ConteudoCompleto);
            }

            // 3. CONSTRUÇÃO E INJEÇÃO DA ESTRUTURA DO BLOCO K HIERÁRQUICA
            Console.WriteLine("      Construindo estrutura hierárquica do Bloco K (K001 -> K100 -> K200 -> K990)...");

            var blocoKCompleto = new List<string>
            {
                "|K001|0|", // Abertura do Bloco com movimento
                $"|K100|{_dataInicio}|{_dataFim}|" // PAI OBRIGATÓRIO: Período de apuração
            };

            // Injeta os filhos (Estoques)
            blocoKCompleto.AddRange(_novosK200);

            // Encerramento: K001 (1) + K100 (1) + total de K200 + K990 (1)
            int totalLinhasBlocoK = 1 + 1 + _novosK200.Count + 1;
            blocoKCompleto.Add($"|K990|{totalLinhasBlocoK}|");

            // 4. POSICIONAMENTO CIRÚRGICO NO ARQUIVO FINAL
            // O Bloco K deve entrar exatamente antes do Bloco 1 ou Bloco 9
            int indiceInsercao = resultadoCru.FindIndex(l =>
                l.Contains("|1001|") || l.StartsWith("1001|") ||
                l.Contains("|9001|") || l.StartsWith("9001|"));

            if (indiceInsercao >= 0)
                resultadoCru.InsertRange(indiceInsercao, blocoKCompleto);
            else
                resultadoCru.AddRange(blocoKCompleto);

            // Sanitização final de delimitadores
            string NormalizarPipes(string linha)
            {
                if (string.IsNullOrWhiteSpace(linha)) return linha;
                linha = linha.Trim();
                if (!linha.StartsWith("|")) linha = "|" + linha;
                if (!linha.EndsWith("|")) linha = linha + "|";
                return linha;
            }

            return resultadoCru.Select(NormalizarPipes).ToList();
        }
    }
}
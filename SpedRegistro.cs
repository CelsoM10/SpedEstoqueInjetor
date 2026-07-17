using System;

namespace SpedEstoqueInjetor
{
    public class SpedRegistro
    {
        public string Registro { get; set; } = string.Empty;
        public string ConteudoCompleto { get; set; } = string.Empty;
        public string[] Campos { get; set; } = Array.Empty<string>();

        public SpedRegistro(string linha)
        {
            // Protege contra nulos e limpa quebras de linha físicas remanescentes
            ConteudoCompleto = linha?.TrimEnd('\r', '\n') ?? string.Empty;
            Campos = ConteudoCompleto.Split('|');

            // Garante a identificação precisa eliminando espaços invisíveis (ex: "0200 " vira "0200")
            if (Campos.Length > 1)
                Registro = Campos[1].Trim();
        }
    }
}
# SPED EFD — Injetor de Estoque (0200 + K200)
## Instruções para Visual Studio Community

---

## O QUE O PROGRAMA FAZ

Todo mes voce coloca a planilha Excel e o arquivo SPED validado numa pasta,
aperta F5 no Visual Studio, e o programa:

1. Le a planilha com os produtos e o estoque
2. Remove os registros 0200 e K200 antigos do SPED
3. Injeta os novos registros corretos
4. Recalcula todos os contadores automaticamente (0990, K990, bloco 9...)
5. Entrega o arquivo SPED corrigido pronto para importar no PVA

---

## INSTALACAO (fazer so uma vez)

### Passo 1 — Instalar o .NET 8 SDK

1. Acesse: https://dotnet.microsoft.com/download/dotnet/8.0
2. Baixe o .NET 8 SDK (x64) para Windows
3. Instale normalmente (avancar, avancar, concluir)

### Passo 2 — Abrir o projeto no Visual Studio Community

1. Abra o Visual Studio Community
2. Clique em "Abrir um projeto ou solucao"
3. Navegue ate a pasta onde salvou estes arquivos e selecione:
       SpedEstoqueInjetor.sln

### Passo 3 — Restaurar pacotes NuGet

1. No menu superior: Projeto -> Restaurar Pacotes NuGet
   (ou aguarde o Visual Studio restaurar automaticamente)
2. O Visual Studio vai baixar o EPPlus automaticamente. Aguarde concluir.

### Passo 4 — Criar a pasta de trabalho

Crie a pasta:
    C:\SPED\Trabalho\

(voce pode mudar esse caminho dentro do codigo — veja secao abaixo)

---

## USO MENSAL (todo mes)

### Passo 1 — Copiar os arquivos para a pasta

Coloque na pasta C:\SPED\Trabalho\ dois arquivos:

    Estoque_Jan2026.xlsx   →  planilha Excel com produtos e estoque
    SPED_Jan2026.txt       →  arquivo SPED ja validado pelo PVA

### Passo 2 — Executar

1. Abra o Visual Studio com o projeto SpedEstoqueInjetor
2. Pressione F5  (ou clique no botao verde Executar)
3. Se houver mais de um Excel ou mais de um .txt na pasta,
   o programa pergunta qual usar — so digitar o numero e Enter

### Passo 3 — Pegar o arquivo gerado

O arquivo corrigido aparece automaticamente na mesma pasta:
    C:\SPED\Trabalho\SPED_Jan2026_CORRIGIDO.txt

Importe esse arquivo no PVA normalmente.

---

## ESTRUTURA DA PLANILHA EXCEL

A planilha deve ter UMA ABA com cabecalho na linha 1:

  Coluna A  →  Codigo Item        ex: SN25
  Coluna B  →  Descricao do Item  ex: SN25 TERMOUMIDOR PARA CHARUTOS...
  Coluna C  →  UND                ex: PC
  Coluna D  →  Tipo Item          ex: 00 - Mercadoria para Revenda
  Coluna E  →  NCM                ex: 8415.82.10
  Coluna F  →  QTD                ex: 197
  Coluna G  →  Tipo Estoque       ex: Estoque de propriedade do informante e em poder de terceiros
  Coluna H  →  CNPJ Terceiro      ex: 30.735.998/0002-14  (vazio se estoque proprio)

IMPORTANTE: o mesmo item pode ter 2 linhas (uma para estoque em poder
de terceiros e outra em poder proprio). O programa trata isso automaticamente.

---

## ALTERAR A PASTA DE TRABALHO

Para mudar a pasta, abra o arquivo Program.cs e altere a linha:

    const string PASTA_TRABALHO = @"C:\SPED\Trabalho";

Exemplo para pasta de rede ou pendrive:

    const string PASTA_TRABALHO = @"D:\Contabilidade\SPED\2026";

---

## PROBLEMAS COMUNS

Programa nao encontra a planilha:
  → Verifique se o arquivo tem extensao .xlsx (nao .xls)

Erro "pacote EPPlus nao encontrado":
  → No Visual Studio: menu Projeto → Restaurar Pacotes NuGet

O PVA ainda da erro apos importar:
  → Envie o relatorio de erros para analise

Quero mudar o codigo da unidade de medida (padrao 13 = PC):
  → Altere no Program.cs:   const string CODIGO_UNID_PC = "13";

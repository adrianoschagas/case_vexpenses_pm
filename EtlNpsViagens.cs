using System.Data;
using System.Text.RegularExpressions;
using ExcelDataReader;

namespace VExpenses.Etl;

// Mapeamento dos campos da planilha cujo a extração de dados vai ser realizada.
// A Classe abaixo será preenchida só na transformação dos dados. 
public record NpsRecord(
    string Id,
    string Mes,
    string Perfil,
    string EtapaParou,
    double? Nota,
    string Comentario)
{
    public string ClasseNps { get; set; } = "";
    public List<string> Temas { get; set; } = new();
    public string EtapaFunilInferida { get; set; } = "";
    public bool SemTema => Temas.Count == 0;
}

// Dicionário de palavras-chave por tema.
public static class Taxonomia
{
    private const RegexOptions Opts = RegexOptions.Compiled | RegexOptions.IgnoreCase;

    public static readonly Dictionary<string, Regex[]> Temas = new()
    {
        ["expiracao_pedido"]  = Build("expir", @"\bprazo\b", "24 hora", "a tempo", "venc", "lembrete"),
        ["visibilidade_fila"] = Build(@"\bfila\b", "quantos pedidos", @"\bordem\b", "urgencia",
                                      "visibilidade", "ingerenciavel", "acompanhar"),
        ["politica_nao_clara"] = Build("politica", @"\bregra\b", "limite por trecho",
                                       "nao estava destacado", "calcular na mao", "dentro ou fora"),
        ["busca_lenta"]        = Build("lent", "demor", "timeout", "carreg", "nunca retorna", "sumiu"),
        ["inventario_pobre"]   = Build("poucas opcoes", "poucos resultado", "2 opcoes",
                                       "pouca opcao", "opcoes de voo", "nem tento"),
        ["preco_fora_mercado"] = Build("caro", "acima do limite", "mais caro",
                                       @"r\$ ?\d", "fora da realidade", "google flights", "preco"),
        ["notificacao_ruim"]   = Build("notificacao", "whatsapp", "alarme", "aviso"),
        ["compra_fora"]        = Build("fora da plataforma", "direto no site", "reservo fora",
                                       "comprando fora", "perdi a concilia", "comprar fora"),
        ["elogio_experiencia"] = Build("simples", "facil", "5 minutos", "tranquilo",
                                       "boa experiencia", "funcionou", "na hora"),
    };

    // Etapas do funil relacionadas ao dicionário de temas.
    public static readonly Dictionary<string, string> TemaParaEtapa = new()
    {
        ["busca_lenta"]        = "busca_selecao",
        ["inventario_pobre"]   = "busca_selecao",
        ["preco_fora_mercado"] = "busca_selecao",
        ["expiracao_pedido"]   = "envio_aprovacao",
        ["visibilidade_fila"]  = "envio_aprovacao",
        ["politica_nao_clara"] = "envio_aprovacao",
        ["notificacao_ruim"]   = "envio_aprovacao",
        ["compra_fora"]        = "pos_funil_evasao",
        ["elogio_experiencia"] = "conversao_ok",
    };

    // Agrupamentos para resumo executivo
    public static readonly Dictionary<string, string[]> MacroDor = new()
    {
        ["aprovacao"]    = new[] { "expiracao_pedido", "visibilidade_fila",
                                   "politica_nao_clara", "notificacao_ruim" },
        ["busca_oferta"] = new[] { "preco_fora_mercado", "busca_lenta", "inventario_pobre" },
    };

    private static Regex[] Build(params string[] patterns) =>
        patterns.Select(p => new Regex(p, Opts)).ToArray();
}

public static class TextOps
{
    // Remoção de acentos e letras minusculas e espaços.
    public static string Normalize(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return "";
        var lowered = s.ToLowerInvariant();
        var noAccents = string.Concat(lowered.Normalize(System.Text.NormalizationForm.FormD)
            .Where(c => System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c)
                        != System.Globalization.UnicodeCategory.NonSpacingMark));
        return Regex.Replace(noAccents, @"\s+", " ").Trim();
    }

    // Classificação do NPS.
    public static string ClassifyNps(double? nota) => nota switch
    {
        null => "sem_nota",
        >= 9 => "promotor",
        >= 7 => "neutro",
        _    => "detrator",
    };
}

public class NpsEtlPipeline
{
    // Lê a planilha e devolve os registros da aba "NPS e Comentários"
    public IEnumerable<NpsRecord> Extract(string caminhoXlsx)
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        using var stream = File.Open(caminhoXlsx, FileMode.Open, FileAccess.Read);
        using var reader = ExcelReaderFactory.CreateReader(stream);

        var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
        {
            ConfigureDataTable = _ => new ExcelDataTableConfiguration
            {
                UseHeaderRow = false  
            }
        });

        // Mapeamento da aba do excel
        var tabela = dataSet.Tables[1];
        var registros = new List<NpsRecord>();

        // Desconsidera as primeiras 4 linhas são título, descrição, linha vazia e cabeçalho
        for (int i = 4; i < tabela.Rows.Count; i++)
        {
            var row = tabela.Rows[i];

            var id = row[0]?.ToString() ?? "";
            if (string.IsNullOrWhiteSpace(id)) continue;

            double? nota = double.TryParse(row[4]?.ToString(), out var n) ? n : null;

            registros.Add(new NpsRecord(
                Id:          id,
                Mes:         row[1]?.ToString() ?? "",
                Perfil:      row[2]?.ToString() ?? "",
                EtapaParou:  row[3]?.ToString() ?? "",
                Nota:        nota,
                Comentario:  row[5]?.ToString() ?? ""
            ));
        }

        return registros;
    }

    // Classifica de cada registro com base em qual nota NPS, quais temas aparecem no comentário
    public List<NpsRecord> Transform(IEnumerable<NpsRecord> registros)
    {

        var prioridadeEtapa = new[]
            { "envio_aprovacao", "busca_selecao", "pos_funil_evasao", "conversao_ok" };

        return registros.Select(r =>
        {
            var norm = TextOps.Normalize(r.Comentario);

            r.ClasseNps = TextOps.ClassifyNps(r.Nota);

            r.Temas = Taxonomia.Temas
                .Where(kv => kv.Value.Any(rgx => rgx.IsMatch(norm)))
                .Select(kv => kv.Key)
                .ToList();
s
            var etapas = r.Temas
                .Select(t => Taxonomia.TemaParaEtapa.GetValueOrDefault(t))
                .Where(e => e is not null)
                .ToHashSet();
            r.EtapaFunilInferida = prioridadeEtapa.FirstOrDefault(etapas.Contains) ?? "indefinido";

            return r;
        }).ToList();
    }

    // Monta as análises e cospe o resultado no console com a construção do csv
    public AnalysisMarts BuildMarts(List<NpsRecord> df)
    {
        var detratores = df.Where(r => r.ClasseNps == "detrator").ToList();
        int nDet = detratores.Count;

        // agrupamento por tema 
        var temasDetratores = detratores
            .SelectMany(r => r.Temas)
            .GroupBy(t => t)
            .Select(g => new ThemeStat(
                Tema: g.Key,
                DetratoresAfetados: g.Count(),
                PctDetratores: nDet == 0 ? 0 : Math.Round(100.0 * g.Count() / nDet, 1)))
            .OrderByDescending(x => x.DetratoresAfetados)
            .ToList();

        // NPS = (promotores - detratores) / total, por mês
        var npsMensal = df
            .GroupBy(r => r.Mes)
            .Select(g =>
            {
                int n = g.Count();
                int p = g.Count(x => x.ClasseNps == "promotor");
                int d = g.Count(x => x.ClasseNps == "detrator");
                return new MonthlyNps(g.Key, n == 0 ? 0 : Math.Round(100.0 * (p - d) / n, 1));
            })
            .ToList();

        // onde no funil os detratores estão travando
        var dorPorEtapa = detratores
            .GroupBy(r => r.EtapaFunilInferida)
            .Select(g => new StageStat(g.Key, g.Select(x => x.Id).Distinct().Count()))
            .OrderByDescending(x => x.Detratores)
            .ToList();

        // comentários que não encaixaram em nenhum tema 
        var filaRevisao = df
            .Where(r => r.SemTema && !string.IsNullOrEmpty(r.Comentario))
            .Select(r => new ReviewItem(r.Id, r.Mes, r.Comentario))
            .ToList();

        return new AnalysisMarts(df, temasDetratores, npsMensal, dorPorEtapa, filaRevisao);
    }

    // Cria o CSV na pasta de saída com a análise
    public void Load(AnalysisMarts marts, string outDir)
    {
        Directory.CreateDirectory(outDir);

        var linhasTemas = new List<string> { "Tema,DetratoresAfetados,PctDetratores" };
        linhasTemas.AddRange(marts.TemasDetratores
            .Select(t => $"{t.Tema},{t.DetratoresAfetados},{t.PctDetratores}"));
        File.WriteAllLines(Path.Combine(outDir, "temas_detratores.csv"), linhasTemas);

        var linhasNps = new List<string> { "Mes,NPS" };
        linhasNps.AddRange(marts.NpsMensal
            .Select(m => $"{m.Mes},{m.Nps}"));
        File.WriteAllLines(Path.Combine(outDir, "nps_mensal.csv"), linhasNps);

        var linhasEtapa = new List<string> { "Etapa,Detratores" };
        linhasEtapa.AddRange(marts.DorPorEtapa
            .Select(e => $"{e.Etapa},{e.Detratores}"));
        File.WriteAllLines(Path.Combine(outDir, "dor_por_etapa.csv"), linhasEtapa);

        var linhasRevisao = new List<string> { "Id,Mes,Comentario" };
        linhasRevisao.AddRange(marts.FilaRevisao
            .Select(r => $"{r.Id},{r.Mes},\"{r.Comentario}\""));
        File.WriteAllLines(Path.Combine(outDir, "fila_revisao.csv"), linhasRevisao);

        Console.WriteLine($"\n[LOAD] 4 arquivos CSV gravados em: {Path.GetFullPath(outDir)}");
    }
}

// Records de saída
public record ThemeStat(string Tema, int DetratoresAfetados, double PctDetratores);
public record MonthlyNps(string Mes, double Nps);
public record StageStat(string Etapa, int Detratores);
public record ReviewItem(string Id, string Mes, string Comentario);
public record AnalysisMarts(
    List<NpsRecord> Fato,
    List<ThemeStat> TemasDetratores,
    List<MonthlyNps> NpsMensal,
    List<StageStat> DorPorEtapa,
    List<ReviewItem> FilaRevisao);

public static class Program
{
    public static void Main()
    {
        var pipeline = new NpsEtlPipeline();

        var brutos = pipeline.Extract("case_vexpenses_dados.xlsx");
        Console.WriteLine($"[EXTRACT] registros lidos");

        var transformados = pipeline.Transform(brutos);
        double cobertura = transformados.Count == 0 ? 0 :
            100.0 * transformados.Count(r => !r.SemTema) / transformados.Count;
        Console.WriteLine($"[TRANSFORM] cobertura de classificação automática: {cobertura:F0}%");

        var marts = pipeline.BuildMarts(transformados);

        Console.WriteLine("\n=== TEMAS ENTRE DETRATORES ===");
        foreach (var t in marts.TemasDetratores)
            Console.WriteLine($"{t.Tema,-22} {t.DetratoresAfetados,3}  ({t.PctDetratores,5:F1}%)");

        Console.WriteLine("\n=== NPS POR MES ===");
        foreach (var m in marts.NpsMensal)
            Console.WriteLine($"{m.Mes,-8}  NPS: {m.Nps,6:F1}");

        Console.WriteLine("\n=== DOR POR ETAPA DO FUNIL ===");
        foreach (var e in marts.DorPorEtapa)
            Console.WriteLine($"{e.Etapa,-22}  {e.Detratores} detratores");

        if (marts.FilaRevisao.Count > 0)
        {
            Console.WriteLine($"\n=== FILA DE REVISAO ({marts.FilaRevisao.Count} comentarios sem tema) ===");
            foreach (var r in marts.FilaRevisao)
                Console.WriteLine($"  [{r.Id}] {r.Comentario}");
        }

        pipeline.Load(marts, "etl_out");

        Console.WriteLine("\nPressione qualquer tecla para fechar...");
        Console.ReadKey();
    }
}

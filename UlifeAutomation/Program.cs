using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using ClosedXML.Excel;
using Microsoft.Extensions.Configuration;
using Oracle.ManagedDataAccess.Client;

namespace UlifeAutomation;

public static class Program
{
    public static async Task<int> Main(string[] args)
    {
        var builder = new ConfigurationBuilder()
            .SetBasePath(AppContext.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .AddEnvironmentVariables(prefix: "ULIFE_");

        var configuration = builder.Build();

        var automationConfig = configuration.GetSection("Automation").Get<AutomationConfig>()
                              ?? throw new InvalidOperationException("Missing Automation configuration section.");

        var isDryRun = args.Contains("--dry-run") || automationConfig.DryRun;
        Console.WriteLine($"DryRun: {isDryRun}");

        var outputDir = Path.GetFullPath(automationConfig.Excel.OutputDirectory);
        Directory.CreateDirectory(outputDir);

        using var oracleConnection = new OracleConnection(automationConfig.Database.ConnectionString);
        if (!isDryRun)
        {
            await oracleConnection.OpenAsync();
        }

        var db = new OracleDatabaseService(oracleConnection, isDryRun);

        var runLog = new List<InsertedParamRow>();

        // STEP 1 - Buscar/gerar valores e inserir parametrização
        foreach (var paramJob in automationConfig.Parametrizacoes)
        {
            Console.WriteLine($"Processando parametrização para CODCONC={paramJob.CodConc} com prefixo={paramJob.Prefix}");

            // Gera dois códigos únicos para tipos 6 e 13 (prefixo-######) que não existam na tabela
            var generatedTokens = await db.GenerateUnusedTokensAsync(
                tableName: paramJob.TableName ?? "dbvestib.param_prova_online",
                valueColumn: paramJob.ValueColumn ?? "VAL_PARAM_PROVA_ONLINE",
                prefix: paramJob.Prefix,
                howMany: 2,
                minSuffix: 100000,
                maxSuffix: 999999);

            var tokenForType6 = $"{paramJob.Prefix}-{generatedTokens[0]}";
            var tokenForType13 = $"{paramJob.Prefix}-{generatedTokens[1]}";

            var rows = new List<(int tipo, string valor)>
            {
                (15, paramJob.LoginUsuario ?? "vestib"),
                (16, paramJob.LoginSenha ?? "vestib@ulife9876"),
                (1,  paramJob.Prefix),
                (6,  tokenForType6),
                (13, tokenForType13)
            };

            foreach (var (tipo, valor) in rows)
            {
                // Idempotente: checa se já existe para o CODCONC+TIPO
                var exists = await db.ExistsAsync(
                    $"select 1 from {paramJob.TableName ?? "dbvestib.param_prova_online"} where CODCONC = :codConc and COD_TPO_PARAM_PROVA_ONLINE = :tipo",
                    new OracleParameter(":codConc", paramJob.CodConc),
                    new OracleParameter(":tipo", tipo));

                if (exists)
                {
                    Console.WriteLine($"Já existe registro para CODCONC={paramJob.CodConc} TIPO={tipo}. Pulando inserção.");
                    continue;
                }

                var insertSql = $@"insert into {paramJob.TableName ?? "dbvestib.param_prova_online"}
(COD_PARAM_PROVA_ONLINE, CODCONC, COD_TPO_PARAM_PROVA_ONLINE, VAL_PARAM_PROVA_ONLINE)
values (param_prova_online_s.nextval, :codConc, :tipo, :valor)
returning COD_PARAM_PROVA_ONLINE into :newId";

                var idParam = await db.InsertReturningIdAsync(
                    insertSql,
                    idOutParamName: ":newId",
                    new OracleParameter(":codConc", paramJob.CodConc),
                    new OracleParameter(":tipo", tipo),
                    new OracleParameter(":valor", valor));

                runLog.Add(new InsertedParamRow
                {
                    CodConc = paramJob.CodConc,
                    Tipo = tipo,
                    Valor = valor,
                    IdGerado = idParam
                });

                Console.WriteLine($"Inserido TIPO={tipo}, VALOR={valor}, ID={idParam}.");
            }
        }

        // STEP 2 - Gerar planilha Excel com os parâmetros
        var excelPath = Path.Combine(outputDir, automationConfig.Excel.FileName);
        ExcelExporter.WriteParametrizacao(excelPath, runLog);
        Console.WriteLine($"Planilha gerada em: {excelPath}");

        // STEP 3 - Chamar endpoint de integração (opcional)
        if (automationConfig.Integration?.Enabled == true)
        {
            var httpResult = await IntegrationService.CallEndpointAsync(automationConfig.Integration);
            Console.WriteLine($"Integração HTTP: {(httpResult.Success ? "OK" : "FALHOU")} - Status: {httpResult.StatusCode} - Corpo: {httpResult.BodyPreview}");
        }

        // STEP 4 - Checar integração (query)
        if (!string.IsNullOrWhiteSpace(automationConfig.Verification?.QuerySql))
        {
            Console.WriteLine("Executando verificação de integração...");
            var resultTable = await db.QueryAsTextTableAsync(automationConfig.Verification.QuerySql);
            Console.WriteLine(resultTable);
        }

        // STEP 5 - Configurar horários do botão de teste e prova
        if (automationConfig.Scheduling is not null)
        {
            var concs = automationConfig.Scheduling.CodConcs?.ToArray() ?? Array.Empty<string>();
            if (concs.Length > 0)
            {
                var examStart = automationConfig.Scheduling.ExamStart;
                var examEnd = automationConfig.Scheduling.ExamEnd;
                var testStart = automationConfig.Scheduling.TestStart;
                var testEnd = automationConfig.Scheduling.TestEnd ?? examStart.AddMinutes(-automationConfig.Scheduling.TestEndOffsetMinutesBeforeExamStart.GetValueOrDefault(10));

                Console.WriteLine($"Atualizando horários para {concs.Length} concursos. Teste: {testStart:dd/MM/yyyy HH:mm:ss} - {testEnd:dd/MM/yyyy HH:mm:ss}. Prova: {examStart:dd/MM/yyyy HH:mm:ss} - {examEnd:dd/MM/yyyy HH:mm:ss}.");

                var placeholders = string.Join(",", concs.Select((_, i) => $":c{i}"));

                var updateTestStartSql = $"update CONCURSO set DAT_START_PROVA_ONLINE_HML = :dt where CODCONC in ({placeholders})";
                var updateTestEndSql = $"update CONCURSO set DAT_FIM_PROVA_ONLINE_HML = :dt where CODCONC in ({placeholders})";
                var updateExamStartSql = $"update CONCURSO set DAT_START_PROVA_ONLINE = :dt where CODCONC in ({placeholders})";
                var updateExamEndSql = $"update CONCURSO set DAT_FIM_PROVA_ONLINE = :dt where CODCONC in ({placeholders})";

                await db.ExecuteWithDateForConcsAsync(updateTestStartSql, testStart, concs);
                await db.ExecuteWithDateForConcsAsync(updateTestEndSql, testEnd, concs);
                await db.ExecuteWithDateForConcsAsync(updateExamStartSql, examStart, concs);
                await db.ExecuteWithDateForConcsAsync(updateExamEndSql, examEnd, concs);
            }
        }

        Console.WriteLine("Rotina finalizada.");
        return 0;
    }
}

public sealed class OracleDatabaseService
{
    private readonly OracleConnection _connection;
    private readonly bool _dryRun;

    public OracleDatabaseService(OracleConnection connection, bool dryRun)
    {
        _connection = connection;
        _dryRun = dryRun;
    }

    public async Task<bool> ExistsAsync(string sql, params OracleParameter[] parameters)
    {
        using var cmd = new OracleCommand(sql, _connection);
        foreach (var p in parameters) cmd.Parameters.Add(p);
        if (_dryRun) return false;
        using var reader = await cmd.ExecuteReaderAsync();
        return await reader.ReadAsync();
    }

    public async Task<long> InsertReturningIdAsync(string sql, string idOutParamName, params OracleParameter[] parameters)
    {
        using var cmd = new OracleCommand(sql, _connection);
        foreach (var p in parameters) cmd.Parameters.Add(p);

        var idParam = new OracleParameter(idOutParamName, OracleDbType.Int64)
        {
            Direction = System.Data.ParameterDirection.Output
        };
        cmd.Parameters.Add(idParam);

        if (_dryRun)
        {
            Console.WriteLine($"[DRY-RUN] {sql.Replace('\n', ' ')}");
            return -1;
        }

        await cmd.ExecuteNonQueryAsync();
        var idVal = Convert.ToInt64(idParam.Value);
        return idVal;
    }

    public async Task<List<int>> GenerateUnusedTokensAsync(string tableName, string valueColumn, string prefix, int howMany, int minSuffix, int maxSuffix)
    {
        var rng = new Random();
        var generated = new HashSet<int>();

        while (generated.Count < howMany)
        {
            var candidate = rng.Next(minSuffix, maxSuffix + 1);
            var composed = $"{prefix}-{candidate}";
            var exists = await ExistsAsync($"select 1 from {tableName} where {valueColumn} = :v", new OracleParameter(":v", composed));
            if (!exists && !generated.Contains(candidate))
            {
                generated.Add(candidate);
            }
        }

        return generated.ToList();
    }

    public async Task<string> QueryAsTextTableAsync(string sql)
    {
        using var cmd = new OracleCommand(sql, _connection);
        if (_dryRun)
        {
            return "[DRY-RUN] Consulta não executada.";
        }

        using var reader = await cmd.ExecuteReaderAsync();
        var sb = new StringBuilder();

        // Header
        for (int i = 0; i < reader.FieldCount; i++)
        {
            if (i > 0) sb.Append(" | ");
            sb.Append(reader.GetName(i));
        }
        sb.AppendLine();
        sb.AppendLine(new string('-', Math.Max(10, sb.Length)));

        // Rows
        while (await reader.ReadAsync())
        {
            for (int i = 0; i < reader.FieldCount; i++)
            {
                if (i > 0) sb.Append(" | ");
                var val = reader.IsDBNull(i) ? "<null>" : reader.GetValue(i)?.ToString();
                sb.Append(val);
            }
            sb.AppendLine();
        }

        return sb.ToString();
    }

    public async Task ExecuteWithDateForConcsAsync(string sqlWithDateAndConcPlaceholders, DateTime dateValue, string[] concs)
    {
        using var cmd = new OracleCommand(sqlWithDateAndConcPlaceholders, _connection);
        cmd.Parameters.Add(":dt", OracleDbType.Date).Value = dateValue;
        for (int i = 0; i < concs.Length; i++)
        {
            cmd.Parameters.Add($":c{i}", OracleDbType.Varchar2).Value = concs[i];
        }

        if (_dryRun)
        {
            Console.WriteLine($"[DRY-RUN] {sqlWithDateAndConcPlaceholders.Replace('\n', ' ')} - dt={dateValue:dd/MM/yyyy HH:mm:ss} - concs=[{string.Join(",", concs)}]");
            return;
        }

        await cmd.ExecuteNonQueryAsync();
    }
}

public static class ExcelExporter
{
    public static void WriteParametrizacao(string filePath, IEnumerable<InsertedParamRow> rows)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Parametrizacao");

        var headers = new[] { "COD_PARAM_ID", "CODCONC", "TIPO", "VALOR" };
        for (int i = 0; i < headers.Length; i++)
        {
            ws.Cell(1, i + 1).Value = headers[i];
        }

        int r = 2;
        foreach (var row in rows.OrderBy(x => x.CodConc).ThenBy(x => x.Tipo))
        {
            ws.Cell(r, 1).Value = row.IdGerado;
            ws.Cell(r, 2).Value = row.CodConc;
            ws.Cell(r, 3).Value = row.Tipo;
            ws.Cell(r, 4).Value = row.Valor;
            r++;
        }

        ws.Columns().AdjustToContents();
        wb.SaveAs(filePath);
    }
}

public static class IntegrationService
{
    public sealed record HttpCallResult(bool Success, int StatusCode, string BodyPreview);

    public static async Task<HttpCallResult> CallEndpointAsync(IntegrationConfig cfg)
    {
        using var http = new HttpClient();
        if (cfg.Headers is not null)
        {
            foreach (var kv in cfg.Headers)
            {
                if (kv.Key.Equals("Authorization", StringComparison.OrdinalIgnoreCase))
                {
                    http.DefaultRequestHeaders.Authorization = AuthenticationHeaderValue.Parse(kv.Value);
                }
                else
                {
                    http.DefaultRequestHeaders.Add(kv.Key, kv.Value);
                }
            }
        }

        using var req = new HttpRequestMessage(new HttpMethod(cfg.Method ?? "POST"), cfg.Url);
        if (!string.IsNullOrWhiteSpace(cfg.BodyJson))
        {
            req.Content = new StringContent(cfg.BodyJson, Encoding.UTF8, "application/json");
        }

        try
        {
            var resp = await http.SendAsync(req);
            var body = await resp.Content.ReadAsStringAsync();
            var preview = body.Length > 400 ? body[..400] + "..." : body;
            return new HttpCallResult(resp.IsSuccessStatusCode, (int)resp.StatusCode, preview);
        }
        catch (Exception ex)
        {
            return new HttpCallResult(false, -1, ex.Message);
        }
    }
}

public sealed class InsertedParamRow
{
    public long IdGerado { get; set; }
    public string CodConc { get; set; } = string.Empty;
    public int Tipo { get; set; }
    public string Valor { get; set; } = string.Empty;
}

public sealed class AutomationConfig
{
    public bool DryRun { get; set; }
    public DatabaseConfig Database { get; set; } = new();
    public List<ParametrizacaoJob> Parametrizacoes { get; set; } = new();
    public IntegrationConfig? Integration { get; set; }
    public VerificationConfig? Verification { get; set; }
    public SchedulingConfig? Scheduling { get; set; }
    public ExcelConfig Excel { get; set; } = new();
}

public sealed class DatabaseConfig
{
    public string ConnectionString { get; set; } = string.Empty;
}

public sealed class ParametrizacaoJob
{
    public string CodConc { get; set; } = string.Empty;
    public string Prefix { get; set; } = string.Empty;
    public string? TableName { get; set; }
    public string? ValueColumn { get; set; }
    public string? LoginUsuario { get; set; }
    public string? LoginSenha { get; set; }
}

public sealed class IntegrationConfig
{
    public bool Enabled { get; set; }
    public string Url { get; set; } = string.Empty;
    public string Method { get; set; } = "POST";
    public Dictionary<string, string>? Headers { get; set; }
    public string? BodyJson { get; set; }
}

public sealed class VerificationConfig
{
    public string? QuerySql { get; set; }
}

public sealed class SchedulingConfig
{
    public List<string>? CodConcs { get; set; }
    public DateTime ExamStart { get; set; }
    public DateTime ExamEnd { get; set; }
    public DateTime TestStart { get; set; }
    public DateTime? TestEnd { get; set; }
    public int? TestEndOffsetMinutesBeforeExamStart { get; set; }
}

public sealed class ExcelConfig
{
    public string OutputDirectory { get; set; } = "output";
    public string FileName { get; set; } = "parametrizacao.xlsx";
}

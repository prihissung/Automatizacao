using Microsoft.Extensions.Configuration;
using AutomacaoVestib.App.Automacao;

var builder = new ConfigurationBuilder()
	.SetBasePath(Directory.GetCurrentDirectory())
	.AddJsonFile("appsettings.json", optional: false, reloadOnChange: false)
	.AddEnvironmentVariables();
var configuration = builder.Build();

var settings = new AppSettings();
configuration.Bind(settings);

var cts = new CancellationTokenSource();
Console.CancelKeyPress += (s, e) =>
{
	e.Cancel = true; cts.Cancel();
};

var oracle = new OracleService(settings.Oracle.ConnectionString);
var http = new HttpIntegrationService(new HttpClient());

string codConc = settings.Automation.CodConcForParametrization;
string basePrefix = settings.Automation.BasePrefix;
var candidateSuffixes = settings.Automation.ProposedSuffixes;
bool dryRun = settings.Automation.DryRun;

try
{
	await oracle.OpenAsync(cts.Token);

	// 1) Buscar nextval não utilizado e inserir parametrização do banco
	Console.WriteLine("[1] Verificando sufíxos de nextval propostos...");
	var existing = await oracle.FindExistingValuesAsync(basePrefix, candidateSuffixes, cts.Token);
	var usedValues = existing.Select(r => r.Valor).ToHashSet(StringComparer.OrdinalIgnoreCase);

	string type6Value;
	string type13Value;
	if (settings.Automation.TryFindUnusedSuffixAutomatically)
	{
		long? s1 = candidateSuffixes.FirstOrDefault(s => !usedValues.Contains($"{basePrefix}-{s}"));
		long? s2 = candidateSuffixes.LastOrDefault(s => !usedValues.Contains($"{basePrefix}-{s}"));
		if (s1 == 0 || s2 == 0 || s1 == s2)
		{
			throw new InvalidOperationException("Não foi possível encontrar dois sufíxos livres distintos nos propostos.");
		}
		type6Value = $"{basePrefix}-{s2}";
		type13Value = $"{basePrefix}-{s1}";
	}
	else
	{
		// Usa exatamente os dois primeiros fornecidos
		type6Value = $"{basePrefix}-{candidateSuffixes.ElementAt(1)}";
		type13Value = $"{basePrefix}-{candidateSuffixes.ElementAt(0)}";
	}
	Console.WriteLine($"Selecionados: tipo6={type6Value}, tipo13={type13Value}");

	var values = new GeneratedValues(
		Type6Value: type6Value,
		Type13Value: type13Value,
		Type1Value: basePrefix,
		LoginValue: settings.Automation.LoginValue,
		EmailValue: settings.Automation.EmailValue
	);

	Console.WriteLine("Inserindo parametrização (tipos 15,16,1,6,13)..." + (dryRun ? " [DRY-RUN]" : string.Empty));
	await oracle.InsertParamAsync(codConc, 15, values.LoginValue, dryRun, cts.Token);
	await oracle.InsertParamAsync(codConc, 16, values.EmailValue, dryRun, cts.Token);
	await oracle.InsertParamAsync(codConc, 1, values.Type1Value, dryRun, cts.Token);
	await oracle.InsertParamAsync(codConc, 6, values.Type6Value, dryRun, cts.Token);
	await oracle.InsertParamAsync(codConc, 13, values.Type13Value, dryRun, cts.Token);

	// 2) Preencher planilha Excel
	Console.WriteLine("[2] Gerando planilha do ULIFE...");
	ExcelService.GenerateUlifeSheet(settings.Automation.ExcelOutputPath, codConc, values);
	Console.WriteLine($"Planilha gerada em: {settings.Automation.ExcelOutputPath}");

	// 3) Rodar endpoint para integrar candidatos
	if (settings.Automation.EnableHttpIntegration)
	{
		Console.WriteLine("[3] Enviando integração HTTP...");
		var resp = await http.SendAsync(settings.Integration, codConc, values, cts.Token);
		Console.WriteLine($"HTTP {resp.StatusCode}");
	}
	else
	{
		Console.WriteLine("[3] Integração HTTP desabilitada por configuração.");
	}

	// 4) Checar integração com query de validação (se fornecida)
	if (!string.IsNullOrWhiteSpace(settings.Automation.ValidationSql))
	{
		Console.WriteLine("[4] Executando query de validação...");
		var table = await oracle.RunArbitraryQueryAsync(settings.Automation.ValidationSql, cts.Token);
		Console.WriteLine($"Linhas retornadas: {table.Rows.Count}");
	}
	else
	{
		Console.WriteLine("[4] Nenhuma query de validação configurada.");
	}

	// 5) Configurar horários do botão de teste e da prova
	Console.WriteLine("[5] Atualizando horários de teste (HML) e prova...");
	var cods = settings.Automation.CodConcForSchedule;
	await oracle.UpdateScheduleAsync(cods, settings.Automation.Schedule.TestStart, settings.Automation.Schedule.TestEnd, homolog: true, dryRun: dryRun, cts.Token);
	await oracle.UpdateScheduleAsync(cods, settings.Automation.Schedule.ExamStart, settings.Automation.Schedule.ExamEnd, homolog: false, dryRun: dryRun, cts.Token);

	Console.WriteLine("Concluído.");
}
catch (Exception ex)
{
	Console.Error.WriteLine($"Erro: {ex.Message}\n{ex}");
	Environment.ExitCode = 1;
}
finally
{
	await oracle.DisposeAsync();
}

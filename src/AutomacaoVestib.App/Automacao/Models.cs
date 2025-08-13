namespace AutomacaoVestib.App.Automacao;

public class OracleSettings
{
	public string ConnectionString { get; set; } = string.Empty;
}

public class ScheduleConfig
{
	public DateTime TestStart { get; set; }
	public DateTime TestEnd { get; set; }
	public DateTime ExamStart { get; set; }
	public DateTime ExamEnd { get; set; }
}

public class AutomationConfig
{
	public string CodConcForParametrization { get; set; } = string.Empty;
	public List<string> CodConcForSchedule { get; set; } = new();
	public string BasePrefix { get; set; } = string.Empty; // ex: "6276"
	public string LoginValue { get; set; } = string.Empty; // ex: "vestib"
	public string EmailValue { get; set; } = string.Empty; // ex: "vestib@ulife9876"
	public bool TryFindUnusedSuffixAutomatically { get; set; } = true;
	public List<long> ProposedSuffixes { get; set; } = new(); // ex: [114735, 225735]
	public string ExcelOutputPath { get; set; } = string.Empty;
	public bool EnableHttpIntegration { get; set; } = false;
	public bool DryRun { get; set; } = true;
	public ScheduleConfig Schedule { get; set; } = new();
	public string ValidationSql { get; set; } = string.Empty; // optional: query grandona
}

public class IntegrationConfig
{
	public string Url { get; set; } = string.Empty;
	public string Method { get; set; } = "POST";
	public Dictionary<string, string> Headers { get; set; } = new();
	public string PayloadTemplate { get; set; } = string.Empty;
}

public class AppSettings
{
	public OracleSettings Oracle { get; set; } = new();
	public AutomationConfig Automation { get; set; } = new();
	public IntegrationConfig Integration { get; set; } = new();
}

public record ParamProvaOnlineRow(decimal CodParam, string CodConc, int Tipo, string Valor);

public record GeneratedValues(string Type6Value, string Type13Value, string Type1Value, string LoginValue, string EmailValue);
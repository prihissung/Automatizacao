using Oracle.ManagedDataAccess.Client;
using System.Data;
using Microsoft.Extensions.Configuration;

namespace AutomacaoVestib.App.Automacao;

public class OracleService : IAsyncDisposable
{
	private readonly string _connectionString;
	private OracleConnection? _conn;

	public OracleService(string connectionString)
	{
		_connectionString = connectionString;
	}

	public async Task OpenAsync(CancellationToken ct)
	{
		_conn = new OracleConnection(_connectionString);
		await _conn.OpenAsync(ct);
	}

	public async Task CloseAsync()
	{
		if (_conn != null)
		{
			await _conn.CloseAsync();
			await _conn.DisposeAsync();
			_conn = null;
		}
	}

	public async ValueTask DisposeAsync()
	{
		await CloseAsync();
	}

	private OracleCommand CreateCommand(string sql)
	{
		if (_conn == null) throw new InvalidOperationException("Connection not opened");
		var cmd = _conn.CreateCommand();
		cmd.BindByName = true;
		cmd.CommandText = sql;
		return cmd;
	}

	public async Task<List<ParamProvaOnlineRow>> FindExistingValuesAsync(string basePrefix, IEnumerable<long> suffixes, CancellationToken ct)
	{
		var list = new List<ParamProvaOnlineRow>();
		var values = suffixes.Select(s => $"{basePrefix}-{s}").ToList();
		string sql = $@"SELECT COD_PARAM_PROVA_ONLINE, CODCONC, COD_TPO_PARAM_PROVA_ONLINE, VAL_PARAM_PROVA_ONLINE
FROM PARAM_PROVA_ONLINE
WHERE VAL_PARAM_PROVA_ONLINE IN ({string.Join(",", values.Select((v, i) => $":v{i}"))})";
		using var cmd = CreateCommand(sql);
		for (int i = 0; i < values.Count; i++)
		{
			cmd.Parameters.Add($":v{i}", OracleDbType.Varchar2, values[i], ParameterDirection.Input);
		}
		using var reader = await cmd.ExecuteReaderAsync(ct);
		while (await reader.ReadAsync(ct))
		{
			list.Add(new ParamProvaOnlineRow(
				reader.GetDecimal(0),
				reader.GetString(1),
				Convert.ToInt32(reader.GetDecimal(2)),
				reader.GetString(3)
			));
		}
		return list;
	}

	public async Task<decimal> GetNextvalAsync(CancellationToken ct)
	{
		string sql = "SELECT param_prova_online_s.NEXTVAL FROM dual";
		using var cmd = CreateCommand(sql);
		var result = await cmd.ExecuteScalarAsync(ct);
		return Convert.ToDecimal(result);
	}

	public async Task<int> InsertParamAsync(string codConc, int tipo, string valor, bool dryRun, CancellationToken ct)
	{
		string sql = @"INSERT INTO dbvestib.param_prova_online (COD_PARAM_PROVA_ONLINE, CODCONC, COD_TPO_PARAM_PROVA_ONLINE, VAL_PARAM_PROVA_ONLINE)
VALUES (param_prova_online_s.nextval, :codconc, :tipo, :valor)";
		if (dryRun)
		{
			return 0;
		}
		using var cmd = CreateCommand(sql);
		cmd.Parameters.Add(":codconc", OracleDbType.Varchar2, codConc, ParameterDirection.Input);
		cmd.Parameters.Add(":tipo", OracleDbType.Int32, tipo, ParameterDirection.Input);
		cmd.Parameters.Add(":valor", OracleDbType.Varchar2, valor, ParameterDirection.Input);
		return await cmd.ExecuteNonQueryAsync(ct);
	}

	public async Task<int> UpdateScheduleAsync(IEnumerable<string> codConcs, DateTime start, DateTime end, bool homolog, bool dryRun, CancellationToken ct)
	{
		string startColumn = homolog ? "DAT_START_PROVA_ONLINE_HML" : "DAT_START_PROVA_ONLINE";
		string endColumn = homolog ? "DAT_FIM_PROVA_ONLINE_HML" : "DAT_FIM_PROVA_ONLINE";
		string sql = $"UPDATE CONCURSO C SET C.{startColumn} = :startDate WHERE C.CODCONC IN ({string.Join(",", codConcs.Select((_, i) => $":c{i}"))});";
		string sql2 = $"UPDATE CONCURSO C SET C.{endColumn} = :endDate WHERE C.CODCONC IN ({string.Join(",", codConcs.Select((_, i) => $":c{i}"))});";
		if (dryRun) return 0;

		using var cmd1 = CreateCommand(sql);
		cmd1.Parameters.Add(":startDate", OracleDbType.Date, start, ParameterDirection.Input);
		int idx = 0; foreach (var c in codConcs) cmd1.Parameters.Add($":c{idx++}", OracleDbType.Varchar2, c, ParameterDirection.Input);
		int rows1 = await cmd1.ExecuteNonQueryAsync(ct);

		using var cmd2 = CreateCommand(sql2);
		cmd2.Parameters.Add(":endDate", OracleDbType.Date, end, ParameterDirection.Input);
		idx = 0; foreach (var c in codConcs) cmd2.Parameters.Add($":c{idx++}", OracleDbType.Varchar2, c, ParameterDirection.Input);
		int rows2 = await cmd2.ExecuteNonQueryAsync(ct);

		return rows1 + rows2;
	}

	public async Task<DataTable> RunArbitraryQueryAsync(string sql, CancellationToken ct)
	{
		using var cmd = CreateCommand(sql);
		using var adapter = new OracleDataAdapter(cmd);
		var table = new DataTable();
		adapter.Fill(table);
		await Task.CompletedTask;
		return table;
	}
}
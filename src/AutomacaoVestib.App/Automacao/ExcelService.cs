using ClosedXML.Excel;

namespace AutomacaoVestib.App.Automacao;

public static class ExcelService
{
	public static void GenerateUlifeSheet(string outputPath, string codConc, GeneratedValues values)
	{
		Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
		using var wb = new XLWorkbook();
		var ws = wb.Worksheets.Add("ULIFE_UPLOAD");
		// Cabeçalho (ajuste conforme necessidade do ULIFE)
		ws.Cell(1, 1).Value = "CODCONC";
		ws.Cell(1, 2).Value = "TIPO";
		ws.Cell(1, 3).Value = "VALOR";

		int row = 2;
		ws.Cell(row, 1).Value = codConc; ws.Cell(row, 2).Value = 15; ws.Cell(row, 3).Value = values.LoginValue; row++;
		ws.Cell(row, 1).Value = codConc; ws.Cell(row, 2).Value = 16; ws.Cell(row, 3).Value = values.EmailValue; row++;
		ws.Cell(row, 1).Value = codConc; ws.Cell(row, 2).Value = 1; ws.Cell(row, 3).Value = values.Type1Value; row++;
		ws.Cell(row, 1).Value = codConc; ws.Cell(row, 2).Value = 6; ws.Cell(row, 3).Value = values.Type6Value; row++;
		ws.Cell(row, 1).Value = codConc; ws.Cell(row, 2).Value = 13; ws.Cell(row, 3).Value = values.Type13Value; row++;

		ws.Columns().AdjustToContents();
		wb.SaveAs(outputPath);
	}
}
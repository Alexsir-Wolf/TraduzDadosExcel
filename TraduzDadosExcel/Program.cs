using Google.Cloud.Translation.V2;
using OfficeOpenXml;

class Program
{
    static void Main(string[] args)
    {
        string chaveAPI = "";
        var path = "";
        var pathSaida = "";

        List<string> textoTraduzido = new List<string>();
        TranslationClient client = TranslationClient.CreateFromApiKey(chaveAPI);
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage(new FileInfo(path)))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; 

            int rowCount = worksheet.Dimension.Rows;
            int colCount = worksheet.Dimension.Columns;
            var count = 0;

            for (int row = 1; row <= rowCount; row++)
            {
                for (int col = 1; col <= colCount; col++)
                {
                    string cellValue = worksheet.Cells[row, col].Text.Trim();
                    var traduzido = TraduzirTexto(client, cellValue, "en");
                    textoTraduzido.Add(traduzido);                  

                    Console.WriteLine($"{count} - Traduzido: {cellValue} para - {traduzido}");
                    count++;
                }
            }

            package.Workbook.Worksheets.Add("Dados Traduzidos (EN)");

            for (int i = 0; i < textoTraduzido.Count; i++)            
                package.Workbook.Worksheets["Dados Traduzidos (EN)"].Cells[i + 1, 1].Value = textoTraduzido[i];
            

            package.SaveAs(new FileInfo(pathSaida));
            Console.WriteLine("Arquivo Excel traduzido com sucesso!");
        }
    }

    static string TraduzirTexto(TranslationClient client, string texto, string idiomaDestino)
    {
        TranslationResult result = client.TranslateText(texto, idiomaDestino);
        return result.TranslatedText.ToUpper().Trim();
    }
}

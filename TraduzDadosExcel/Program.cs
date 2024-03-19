using System.Text.RegularExpressions;
using Google.Cloud.Translation.V2;
using OfficeOpenXml;

class Program
{
    static void Main(string[] args)
    {
        // Definir o contexto de licença do EPPlus para uso não comercial
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        Start();
    }

    /// <summary>
    /// Method responsible for start application
    /// </summary>
    private static void Start()
    {
        Console.WriteLine("O que você deseja fazer? \n\n" +
            "Opção A) Retirar os acentos.\n" +
            "Opção B) Trazuzir texto \n" +
            "Opção Z) Sair \n");

        string selectOptions = Console.ReadLine();
        selectOptions = selectOptions.ToUpper();

        if (string.IsNullOrWhiteSpace(selectOptions))
        {
            Console.WriteLine("Por favor, insira uma opção válida. \n");
            Start();
        }
        else if (!selectOptions.Equals("A", StringComparison.OrdinalIgnoreCase) &&
                 !selectOptions.Equals("B", StringComparison.OrdinalIgnoreCase) &&
                 !selectOptions.Equals("Z", StringComparison.OrdinalIgnoreCase))
        {
            Console.WriteLine($"Por favor, o valor digitado '{selectOptions.ToUpper()}' é inválido, insira um valor conforme orientação em tela. \n");
            Start();
        }
        else if (selectOptions.Equals("A", StringComparison.OrdinalIgnoreCase))
        {
            RemoveAccents();
        }
        else if (selectOptions.Equals("B", StringComparison.OrdinalIgnoreCase))
        {
            TranslateTexts();
        }
        else if (selectOptions.Equals("Z", StringComparison.OrdinalIgnoreCase))
        {
            Environment.Exit(0);
        }
    }

    /// <summary>
    /// Method responsible for translate text of using API Google.
    /// </summary>
    private static void TranslateTexts()
    {
        string chaveAPI = "";
        var path = "";
        var pathSaida = "";

        List<string> textoTraduzido = new List<string>();
        TranslationClient client = TranslationClient.CreateFromApiKey(chaveAPI);

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

    /// <summary>
    /// Method responsible for remove accents of text
    /// </summary>
    private static void RemoveAccents()
    {
        // Path of the input file (spreadsheet)
        string inputFilePath = "C:\\Projects\\Pasta1.xlsx";

        // Path of the output file (new spreadsheet without special characters)
        string outputFilePath = @"C:\\Projects\\planilha_sem_acento.xlsx";

        try
        {
            // Load the input spreadsheet
            FileInfo inputFile = new FileInfo(inputFilePath);

            using (ExcelPackage package = new ExcelPackage(inputFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // Iterate over the cells of the spreadsheet and remove special characters
                for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                {
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        string cellValue = worksheet.Cells[row, col].Text;
                        string cleanValue = RemoveSpecialCharacters(cellValue);
                        worksheet.Cells[row, col].Value = cleanValue;
                    }
                }

                // Save the new spreadsheet without special characters
                FileInfo outputFile = new FileInfo(outputFilePath);
                package.SaveAs(outputFile);
            }

            Console.WriteLine($"Caracteres especiais removidos com sucesso. Nova planilha salva em: {outputFilePath}, bora executar outra atividade?");
            Start();
        }
        catch (Exception ex)
        {
            Console.WriteLine("Ocorreu um erro: " + ex.Message);
        }
    }

    /// <summary>
    /// Tandling regex for text with accents
    /// </summary>
    /// <param name="input">Text for Tandling</param>
    /// <returns>Return text clean</returns>
    static string RemoveSpecialCharacters(string input)
    {
        // Regular expression to find special characters and accents
        string pattern = "[^a-zA-Z0-9 ]";

        // Substituir os caracteres especiais por vazio
        string clean = Regex.Replace(input, pattern, "");

        return clean;
    }

    /// <summary>
    /// Method responsible for translate text
    /// </summary>
    /// <param name="client">Object single of TranslationClient</param>
    /// <param name="texto">Text for translate</param>
    /// <param name="idiomaDestino">Set which language will be translated</param>
    /// <returns>Return text translate.</returns>
    static string TraduzirTexto(TranslationClient client, string texto, string idiomaDestino)
    {
        TranslationResult result = client.TranslateText(texto, idiomaDestino);
        return result.TranslatedText.ToUpper().Trim();
    }
}

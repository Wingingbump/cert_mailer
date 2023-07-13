using OfficeOpenXml;

namespace cert_mailer;
internal class Evaluations
{
    private string path
    {
        get; set;
    }
    public Evaluations(FileInfo evaluation, string EOCpath, string type)
    {
        // Load the evaluation file and get the first worksheet.
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using var evaluationExcel = new ExcelPackage(evaluation);
        var evalSheet = evaluationExcel.Workbook.Worksheets.FirstOrDefault();

        path = EOCpath;
    }
}

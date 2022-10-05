// See https://aka.ms/new-console-template for more information
using OfficeOpenXml;
using Shane32.ExcelLinq.Builders;
using Shane32.ExcelLinq;

public class TestFileContext : ExcelContext
{
    // in order to read files, you'll need one of these constructors
    public TestFileContext(System.IO.Stream stream) : base(stream) { }
    public TestFileContext(string filename) : base(filename) { }
    public TestFileContext(ExcelPackage excelPackage) : base(excelPackage) { }

    // in order to write new files, you'll need a default constructor
    public TestFileContext() : base() { }

    // define an easy way to access the sheet1 table
    public List<RecipeDurations> RecipeDurations => GetSheet<RecipeDurations>();
    public List<RecipeInputs> RecipeInputs => GetSheet<RecipeInputs>();
    public List<RecipeOutputs> RecipeOutputs => GetSheet<RecipeOutputs>();
    public List<Query> Query => GetSheet<Query>("Query");
    public List<QueryResult> Results => GetSheet<QueryResult>();
    public List<Query> Remainder => GetSheet<Query>("Remainder");

    protected override void OnModelCreating(ExcelModelBuilder builder)
    {
        Action<ExcelRange> numberFormatter2 = range => {
            range.Style.Numberformat.Format = "0.00";
            range[1, range.Start.Column].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        };
        Action<ExcelRange> numberFormatter0 = range =>
        {
            range.Style.Numberformat.Format = "0";
            range[1, range.Start.Column].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        };
        Action<ExcelWorksheet, ExcelRange> sheetPolisher = (sheet, range) =>
        {
            range[1, 1, 1, range.Columns].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            sheet.Cells.AutoFitColumns();
            for (int col = 1; col <= sheet.Cells.Columns; col++)
            {
                sheet.Column(col).Width *= 1.2;
            }
        };

        var sheet1 = builder.Sheet<RecipeDurations>();
        sheet1.Column(x => x.DurationHours, "Hours") //excel heading name is hours
            .ColumnFormatter(numberFormatter2);
        sheet1.Column(x => x.DurationTicks, "Duration")
            .ColumnFormatter(numberFormatter0);
        sheet1.Column(x => x.Building);
        sheet1.Column(x => x.RecipeName, "Key");
        sheet1.WritePolisher(sheetPolisher);

        var sheet2 = builder.Sheet<RecipeInputs>();
        sheet2.Column(x => x.Input);             // define a column for the Date property; look for a column with the name "Date"
        sheet2.Column(x => x.RecipeName, "Key");  // for Quantity, look for a column with the name "Qty"
        sheet2.Column(x => x.Amount)        // for Description, look for a column named either "Description" or "Desc"
            .ColumnFormatter(numberFormatter0);
        sheet2.WritePolisher(sheetPolisher);

        var sheet3 = builder.Sheet<RecipeOutputs>();
        sheet3.Column(x => x.Amount)              // define a column for the Date property; look for a column with the name "Date"
            .ColumnFormatter(numberFormatter0);
        sheet3.Column(x => x.RecipeName, "Key");  // for Quantity, look for a column with the name "Qty"
        sheet3.Column(x => x.Output);       // for Description, look for a column named either "Description" or "Desc"
        sheet3.WritePolisher(sheetPolisher);

        var sheet4 = builder.Sheet<Query>();
        sheet4.Column(x => x.Quantity)
            .ColumnFormatter(numberFormatter2);
        sheet4.Column(x => x.Material);
        sheet4.Column(x => x.TimeframeHours)
            .ColumnFormatter(numberFormatter2);
        sheet4.WritePolisher(sheetPolisher);

        var sheet5 = builder.Sheet<QueryResult>();
        sheet5.Optional();
        sheet5.Column(x => x.RecipeName);
        sheet5.Column(x => x.Building);
        sheet5.Column(x => x.Expertise);
        sheet5.Column(x => x.QuantityOfBuildingsRunningRecipe, "Quantity")
            .ColumnFormatter(numberFormatter2);
        sheet5.WritePolisher(sheetPolisher);

        var sheet6 = builder.Sheet<Query>("Remainder");
        sheet6.Optional();
        sheet6.Column(x => x.Quantity)
            .ColumnFormatter(numberFormatter2);
        sheet6.Column(x => x.Material);
        sheet6.Column(x => x.TimeframeHours)
            .ColumnFormatter(numberFormatter2);
        sheet6.WritePolisher(sheetPolisher);
    }
}

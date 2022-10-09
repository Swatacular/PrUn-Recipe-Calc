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

    public List<BuildingExpertise> BuildingExpertise => GetSheet<BuildingExpertise>();

    public List<BuildingWorkforces> BuildingWorkforces => GetSheet<BuildingWorkforces>();
    public List<WorkforceArea> WorkforceArea => GetSheet<WorkforceArea>();
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
            var columns = range.Columns;
            range[1, 1, 1, columns].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            sheet.Cells.AutoFitColumns();
            for (int col = 1; col <= columns; col++)
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
        sheet2.Column(x => x.Input);            
        sheet2.Column(x => x.RecipeName, "Key");  // for RecipeName, look for a column with the name "Key"
        sheet2.Column(x => x.Amount)   
            .ColumnFormatter(numberFormatter0);
        sheet2.WritePolisher(sheetPolisher);

        var sheet3 = builder.Sheet<RecipeOutputs>();
        sheet3.Column(x => x.Amount)           
            .ColumnFormatter(numberFormatter0);
        sheet3.Column(x => x.RecipeName, "Key");
        sheet3.Column(x => x.Output); 
        sheet3.WritePolisher(sheetPolisher);


        var sheet4 = builder.Sheet<BuildingExpertise>();
        sheet4.Column(x => x.Ticker);
        sheet4.Column(x => x.Name);
        sheet4.Column(x => x.Area);
        sheet4.Column(x => x.Expertise);
        sheet4.WritePolisher(sheetPolisher);

        var sheet5 = builder.Sheet<BuildingWorkforces>();
        sheet5.Column(x => x.Key);
        sheet5.Column(x => x.Building);
        sheet5.Column(x => x.Type);
        sheet5.Column(x => x.Capacity);
        sheet5.Column(x => x.Weight); //How much of the production lines productivity is attributed to that workforce
        sheet5.WritePolisher(sheetPolisher);

        var sheet6 = builder.Sheet<WorkforceArea>();
        sheet6.Column(x => x.Type);
        sheet6.Column(x => x.MinAreaPer1);
        sheet6.Column(x => x.AvgAreaPer1);
        sheet6.Column(x => x.MaxAreaPer1);
        sheet6.Column(x => x.HBB); //How much area is needed per 1 worker if provided by HBB, Should be NULL if HBB provides for 0 workers
        sheet6.Column(x => x.HBC); //^ HBC
        sheet6.Column(x => x.HBM);
        sheet6.Column(x => x.HBL);
        sheet6.WritePolisher(sheetPolisher);

        var sheet7 = builder.Sheet<Query>();
        sheet7.Column(x => x.Quantity)
            .ColumnFormatter(numberFormatter2);
        sheet7.Column(x => x.Material);
        sheet7.Column(x => x.TimeframeHours)
            .ColumnFormatter(numberFormatter2);
        sheet7.WritePolisher(sheetPolisher);

        var sheet8 = builder.Sheet<QueryResult>();
        sheet8.Optional();
        sheet8.Column(x => x.RecipeName);
        sheet8.Column(x => x.Building);
        sheet8.Column(x => x.Expertise);
        sheet8.Column(x => x.QuantityOfBuildingsRunningRecipe, "Quantity")
            .ColumnFormatter(numberFormatter2);
        sheet8.WritePolisher(sheetPolisher);

        var sheet9 = builder.Sheet<Query>("Remainder");
        sheet9.Optional();
        sheet9.Column(x => x.Quantity)
            .ColumnFormatter(numberFormatter2);
        sheet9.Column(x => x.Material);
        sheet9.Column(x => x.TimeframeHours)
            .ColumnFormatter(numberFormatter2);
        sheet9.WritePolisher(sheetPolisher);
    }
}

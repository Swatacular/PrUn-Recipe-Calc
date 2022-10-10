// See https://aka.ms/new-console-template for more information

using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

var excelfile = new TestFileContext("C:\\Users\\Tom\\OneDrive\\Prosperous Universe Shipyard.xlsx");


Console.WriteLine("Hello, World!");

var workforceAndAreas = (
    from workforce in excelfile.BuildingWorkforces
    join area in excelfile.WorkforceArea on workforce.Type equals area.Type
    select new { Workforce = workforce, Area = area })
    .ToLookup(x => x.Workforce.Building, StringComparer.OrdinalIgnoreCase);

var buildings = (
    from expertise in excelfile.BuildingExpertise
    select new Building
    (
        expertise.Name,
        expertise.Ticker,
        expertise.Expertise,
        expertise.Area,
        workforceAndAreas[expertise.Ticker]
            .Select(row => new Workforce
            (
                row.Workforce.Type,
                row.Workforce.Capacity,
                row.Workforce.Weight,
                row.Area.MaxAreaPer1 * row.Workforce.Capacity
            ))
            .ToDictionary(x => x.Type, StringComparer.OrdinalIgnoreCase)
    ))
    .ToDictionary(building => building.Ticker, StringComparer.OrdinalIgnoreCase);

var recipes = excelfile.RecipeDurations
    .Select(recipe => new Recipe(
        recipe.RecipeName,
        recipe.DurationHours,
        recipe.DurationTicks,
        buildings[recipe.Building],
        excelfile.RecipeInputs.Where(x => x.RecipeName == recipe.RecipeName).ToList(),
        excelfile.RecipeOutputs.Where(x => x.RecipeName == recipe.RecipeName).ToList()
    ))
    .OrderBy(x => x.Outputs.Count) //Switch to decending to get AML stuff
    .ToList();

// balances by hour
var requirements = excelfile.Query.GroupBy(x => x.Material).Select(x => new Requirement
{
    Material = x.Key,
    QuantityPerHour = x.Select(y => y.Quantity / y.TimeframeHours).Sum(),
}).ToDictionary(x => x.Material);

var targetHours = excelfile.Query.Select(x => x.TimeframeHours).Max();

var queryResults = new Dictionary<Recipe, QueryResult>();

while (true)
{
    var match = requirements.Values
        .Where(x => x.QuantityPerHour > 0.00000000001)
        .Select(x => new { Requirement = x, Recipe = FindRecipe(x.Material) })
        .Where(x => x.Recipe != null)
        .FirstOrDefault();
    if (match == null) break;
    Console.WriteLine($"Adding recipe '{match.Recipe!.RecipeName}' to fulfill requirement of '{match.Requirement.Material}' at {match.Requirement.QuantityPerHour}/hr");
    Console.WriteLine($"  Recipe requires {string.Join(", ", match.Recipe.Inputs.Select(input => $"{input.Amount / match.Recipe.DurationHours} of '{input.Input}'"))}");
    Console.WriteLine($"  Recipe produces {string.Join(", ", match.Recipe.Outputs.Select(output => $"{output.Amount / match.Recipe.DurationHours} of '{output.Output}'"))}");

    // add building to list
    var quantity = match.Requirement.QuantityPerHour / (match.Recipe!.Outputs.Where(y => y.Output == match.Requirement.Material).Select(y => y.Amount).Sum() / match.Recipe.DurationHours);
    Console.WriteLine($"  Adding {quantity} recipes");
    if (queryResults.TryGetValue(match.Recipe, out var result))
    {
        result.QuantityOfBuildingsRunningRecipe += quantity;
    }
    else
    {
        queryResults.Add(match.Recipe, new QueryResult
        {
            RecipeName = match.Recipe.RecipeName,
            Building = match.Recipe.Building,
            QuantityOfBuildingsRunningRecipe = quantity,
        });
    }

    // remove outputs from requirements
    foreach (var output in match.Recipe.Outputs)
    {
        if (output.Output == match.Requirement.Material)
        {
            //Console.WriteLine($"THIS SHOULD BE ZERO: {requirements[output.Output].QuantityPerHour} / {output.Amount / match.Recipe.DurationHours * quantity} / {requirements[output.Output].QuantityPerHour - (output.Amount / match.Recipe.DurationHours * quantity)}");
            requirements.Remove(match.Requirement.Material);
        }
        else if (requirements.ContainsKey(output.Output))
        {
            requirements[output.Output].QuantityPerHour -= output.Amount / match.Recipe.DurationHours * quantity;
        }
        // todo: if zero remove from list (optional)
    }

    // add inputs to requirements
    foreach (var input in match.Recipe.Inputs)
    {
        if (requirements.TryGetValue(input.Input, out var req))
        {
            req.QuantityPerHour += input.Amount / match.Recipe.DurationHours * quantity;
        }
        else
        {
            requirements.Add(input.Input, new Requirement
            {
                Material = input.Input,
                QuantityPerHour = input.Amount / match.Recipe.DurationHours * quantity,
            });
        }
    }
}

foreach (var result in queryResults.Values)
{
    result.BuildingTicker = result.Building.Ticker;
    result.Expertise = result.Building.Expertise;
    result.BuildingArea = result.Building.Area * result.QuantityOfBuildingsRunningRecipe;
    var pioneers = result.Building.Workforce.GetValueOrDefault("PIONEER");
    result.PioneerQuantity = (pioneers?.Capacity ?? 0) * result.QuantityOfBuildingsRunningRecipe;
    result.PioneerArea = (pioneers?.Area ?? 0) * result.QuantityOfBuildingsRunningRecipe;
    var settlers = result.Building.Workforce.GetValueOrDefault("SETTLER");
    result.SettlerQuantity = (settlers?.Capacity ?? 0) * result.QuantityOfBuildingsRunningRecipe;
    result.SettlerArea = (settlers?.Area ?? 0) * result.QuantityOfBuildingsRunningRecipe;
    var technicians = result.Building.Workforce.GetValueOrDefault("TECHNICIAN");
    result.TechnicianQuantity = (technicians?.Capacity ?? 0) * result.QuantityOfBuildingsRunningRecipe;
    result.TechnicianArea = (technicians?.Area ?? 0) * result.QuantityOfBuildingsRunningRecipe;
    var engineers = result.Building.Workforce.GetValueOrDefault("ENGINEER");
    result.EngineerQuantity = (engineers?.Capacity ?? 0) * result.QuantityOfBuildingsRunningRecipe;
    result.EngineerArea = (engineers?.Area ?? 0) * result.QuantityOfBuildingsRunningRecipe;
    var scientists = result.Building.Workforce.GetValueOrDefault("SCIENTIST");
    result.ScientistQuantity = (scientists?.Capacity ?? 0) * result.QuantityOfBuildingsRunningRecipe;
    result.ScientistArea = (scientists?.Area ?? 0) * result.QuantityOfBuildingsRunningRecipe;
}

// write remainder
excelfile.Remainder.Clear();
foreach (var input in requirements)
{
    excelfile.Remainder.Add(new Query
    {
        Material = input.Key,
        Quantity = input.Value.QuantityPerHour, // * targetHours
        TimeframeHours = 1, // targetHours
    });
}

// write query results
excelfile.Results.Clear();
excelfile.Results.AddRange(queryResults.Values.OrderByDescending(x => x.QuantityOfBuildingsRunningRecipe));

Console.WriteLine("Formatting document");

excelfile.SerializeToFile("C:\\Users\\Tom\\OneDrive\\Prosperous Universe Shipyard - output.xlsx");

Console.WriteLine("Done");

Recipe? FindRecipe(string material)
{
    return recipes
        //Makes a list of recipes that produce the material
        .Where(x => x.Outputs.Any(y => y.Output == material))
        //takes the list, and orders it by what recipe produces the most of that material per hour
        .OrderBy(x => x.Outputs.Where(y => y.Output == material).Select(y => y.Amount).Sum() / x.DurationHours)
        //selects the top/first recipe (or returns null if there are no recipes found
        .FirstOrDefault();
}

internal record Recipe(string RecipeName, double DurationHours, int DurationTicks, Building Building, List<RecipeInputs> Inputs, List<RecipeOutputs> Outputs);

// See https://aka.ms/new-console-template for more information

using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

var excelfile = new TestFileContext("C:\\Users\\Tom\\OneDrive\\Prosperous Universe Shipyard.xlsx");


Console.WriteLine("Hello, World!");

var recipes = excelfile.RecipeDurations
    .Select(recipe => new Recipe(
        recipe.RecipeName,
        recipe.DurationHours,
        recipe.DurationTicks,
        recipe.Building,
        excelfile.BuildingExpertise.Where(x => x.Ticker == recipe.Building).FirstOrDefault().Expertise,
        excelfile.RecipeInputs.Where(x => x.RecipeName == recipe.RecipeName).ToList(),
        excelfile.RecipeOutputs.Where(x => x.RecipeName == recipe.RecipeName).ToList()
    ))
    .OrderBy(x => x.Outputs.Count) //Switch to decending to get AML stuff
    .ToList();

// balances by hour
var requirements = excelfile.Query.GroupBy(x => x.Material).Select(x => new Requirement {
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
            Expertise = match.Recipe.Expertise,
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
        //takes the list, and orders it by what recipe produces the most of that material
        .OrderByDescending(x => x.Outputs.Where(y => y.Output == material).Select(y => y.Amount).Sum() / x.DurationHours)
        //selects the top/first recipe (or returns null if there are no recipes found
        .FirstOrDefault();
}

internal record Recipe(string RecipeName, double DurationHours, int DurationTicks, string Building, string Expertise, List<RecipeInputs> Inputs, List<RecipeOutputs> Outputs);

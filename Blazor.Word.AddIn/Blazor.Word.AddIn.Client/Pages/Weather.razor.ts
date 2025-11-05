/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */

/**
 * Basic function to show how to insert a table into a Word Presentation.
 */
console.log("Loading Weather.razor.ts");

export async function createWeatherTable(forecasts: any[]) {

  console.log("We are now entering function: createWeatherTable");
  console.log("Received forecasts:", forecasts);

  try {
    await Word.run(async function (context) {
      // Get the current document body
      const body = context.document.body;

      // Insert a title
      const title = body.insertParagraph("Weather Forecast", Word.InsertLocation.end);
      title.styleBuiltIn = Word.BuiltInStyleName.title;

      // Create a table with headers: Date, Temp (C), Temp (F), Summary
      // Add 1 for header row + number of forecast rows
      const table = body.insertTable(forecasts.length + 1, 4, Word.InsertLocation.end);
      table.styleBuiltIn = Word.BuiltInStyleName.gridTable5Dark_Accent1;

      // Load the table rows before accessing them
      table.rows.load("items");
      await context.sync();

      // Set header row - load cells before accessing
      const headerRow = table.rows.items[0];
      headerRow.cells.load("items");
      await context.sync();
      
      headerRow.cells.items[0].value = "Date";
      headerRow.cells.items[1].value = "Temp (°C)";
      headerRow.cells.items[2].value = "Temp (°F)";
      headerRow.cells.items[3].value = "Summary";

      // Load cells for all rows at once
      table.rows.items.forEach(row => row.cells.load("items"));
      await context.sync();

      // Now populate data rows without additional sync calls
      for (let i = 0; i < forecasts.length; i++) {
        const row = table.rows.items[i + 1];
        const forecast = forecasts[i];
        
        row.cells.items[0].value = forecast.date;
        row.cells.items[1].value = forecast.temperatureC.toString();
        row.cells.items[2].value = forecast.temperatureF.toString();
        row.cells.items[3].value = forecast.summary || "";
      }

      // Add some spacing after the table
      body.insertParagraph("", Word.InsertLocation.end);

      // Single sync at the end
      await context.sync();
      console.log("Weather Forecast table created successfully.");
    });
  } catch (error) {
    console.error("Error creating Weather Forecast table: ", error);
  }
}
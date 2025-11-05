/*
 * Copyright (c) Maarten van Stam. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

console.log("Loading command.js");

/**
 * Function to run Office JavaScript without any Interop.
 * @param event
 */
async function insertTextInWord(event: Office.AddinCommands.Event) {

  console.log("In insertTextInWord");

  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.insertText("Hello World", Word.InsertLocation.end);
      await context.sync();
    });
  } catch (error: any) {
    console.error(error);
  }
  finally {
    console.log("Finish insertTextInWord");
  }

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

/**
 * Writes the text from the Home Blazor Page to the Word slide
 * @param {any} event
 */
async function callBlazorOnHome(event: Office.AddinCommands.Event) {

  // Implement your custom code here. The following code is a simple Word example.  
  try {
    console.log("Running callBlazorOnHome");

    console.log("Before callStaticLocalComponentMethodinit");
    await callStaticLocalComponentMethodinit("SayHelloHome");
    console.log("After callStaticLocalComponentMethodinit");
  }
  catch (error: any) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    console.error(error);
  }
  finally {
    console.log("Finish callBlazorOnHome");
  }

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}

/**
 * Writes the text from the Counter Blazor Page to the Word slide
 * @param {any} event
 */
async function callBlazorOnCounter(event: Office.AddinCommands.Event) {

  // Implement your custom code here. The following code is a simple Word example.  
  try {
    console.log("Running callBlazorOnHome");

    console.log("Before callStaticLocalComponentMethodinit");
    await callStaticLocalComponentMethodinit("SayHelloCounter");
    console.log("After callStaticLocalComponentMethodinit");
  }
  catch (error: any) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    console.error(error);
  }
  finally {
    console.log("Finish callBlazorOnCounter");
  }

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}

/**
 * Checks if the .NET runtime is loaded and invokes a .NET method to retrieve a string.
 * The string is then inserted into a Word slide as a text box.
 * and some format is added to the text box.
 * 
 * @param {string} methodname - The name of the .NET method to invoke.
 */
async function callStaticLocalComponentMethodinit(methodname: string) {
  console.log("In callStaticLocalComponentMethodinit");

  try {
    let name = "Initializing";

    try {

      var dotnetloaded = await preloadDotNet();

      if (dotnetloaded === true) {
        // Call JSInvokable Function here ...
        name = await DotNet.invokeMethodAsync(
          "Blazor.Word.AddIn.Client",
          methodname,
          "Blazor Fan");
      }
      else {
        name = "Init DotNet Failed";
      }
    } catch (error: any) {
      name = error.message;
      console.error("Error during DotNet invocation: " + name);
    }

    await Word.run(async (context) => {
      const body = context.document.body;
      body.insertText(name, Word.InsertLocation.end);
      await context.sync();
    });

    console.log("Finished Initializing: " + name)
  }
  catch (error: any) {
    console.error(error);
  }
  finally {
    console.log("Finish callStaticLocalComponentMethodinit");
  }
}

/**
 * Local function to preload the .NET runtime and ensure it is ready for use.
 *
 * This function attempts to invoke a dummy method in the .NET runtime to check if it is loaded.
 * It retries up to 5 times with a 1-second delay between attempts if the runtime is not loaded.
 *
 * This won't be necessary if the task pane is automatically opened when the add-in is loaded.
 * Also feel it should be possible to preload in the module.ts file for a guaranteed load.
 *
 * @returns result - Returns true if the .NET runtime is successfully loaded, otherwise false.
 */
async function preloadDotNet() {

  console.log("In preloadDotNet");

  try {
    console.log("Before DotNet.invokeMethodAsync");
    var result = "";

    let attempts = 0;
    while (result === "" && attempts < 5) {
      try {
        if (attempts > 0) {
          await new Promise((resolve) => setTimeout(resolve, 1000));
        }
        result = await DotNet.invokeMethodAsync(
          "Blazor.Word.AddIn.Client",
          "PreloaderDummy"
        );
      } catch (error: any) {
        console.error("Error during DotNet invocation: " + error.message);
      }
      attempts++;
    }

    console.log("After DotNet.invokeMethodAsync");
    return result === "Loaded" ? true : false;

  } catch (error: any) {

    console.log("Error call : " + error.message);
    return false;

  } finally {

    console.log("Finish preloadDotNet");
  }
}

// Associate the functions with their named counterparts in the manifest XML.
Office.actions.associate("insertTextInWord", insertTextInWord);
Office.actions.associate("callBlazorOnHome", callBlazorOnHome);
Office.actions.associate("callBlazorOnCounter", callBlazorOnCounter);
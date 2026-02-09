/**
 * Copyright (c) Maarten van Stam. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
console.log("Loading command.js");

/**
 * Inserts "Hello World" text at the end of the Word document body.
 * This function demonstrates basic Office JavaScript API usage without Blazor interop.
 * 
 * @param event - The Office add-in command event object
 * @returns A promise that resolves when the text insertion is complete
 */
async function insertTextInWord(event: Office.AddinCommands.Event): Promise<void> {
  console.log("In insertTextInWord");

  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.insertText("Hello World from TypeScript", Word.InsertLocation.end);
      await context.sync();
    });
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    console.error("Error in insertTextInWord:", errorMessage);
  } finally {
    console.log("Finish insertTextInWord");
  }

  // Be sure to indicate when the add-in command function is complete
  if (event && typeof event.completed === 'function') {
    event.completed();
  }
}

/**
 * Writes the text from the Home Blazor Page to the Word document.
 * This function invokes a .NET method through Blazor interop to retrieve content
 * from the Home page and insert it into the Word document.
 * 
 * @param event - The Office add-in command event object
 * @returns A promise that resolves when the operation is complete
 */
async function callBlazorOnHome(event: Office.AddinCommands.Event): Promise<void> {

  try {

    console.log("Running callBlazorOnHome");
    await callStaticLocalComponentMethodinit("SayHelloHome");

  } catch (error) {

    const errorMessage = error instanceof Error ? error.message : String(error);
    console.error("Error in callBlazorOnHome:", errorMessage);

  } finally {

    console.log("Finish callBlazorOnHome");
  }

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  if (event && typeof event.completed === 'function') {
    event.completed();
  }
}

/**
 * Writes the text from the Counter Blazor Page to the Word document.
 * This function invokes a .NET method through Blazor interop to retrieve content
 * from the Counter page and insert it into the Word document.
 * 
 * @param event - The Office add-in command event object
 * @returns A promise that resolves when the operation is complete
 */
async function callBlazorOnCounter(event: Office.AddinCommands.Event): Promise<void> {
  try {

    console.log("Running callBlazorOnCounter");
    await callStaticLocalComponentMethodinit("SayHelloCounter");

  } catch (error) {

    const errorMessage = error instanceof Error ? error.message : String(error);
    console.error("Error in callBlazorOnCounter:", errorMessage);

  } finally {

    console.log("Finish callBlazorOnCounter");
  }

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  if (event && typeof event.completed === 'function') {
    event.completed();
  }
}

/**
 * Checks if the .NET runtime is loaded and invokes a .NET method to retrieve a string.
 * The string is then inserted into the Word document body at the end.
 * 
 * @param {string} methodname - The name of the .NET method to invoke.
 * @returns {Promise<void>} A promise that resolves when the operation is complete.
 */
async function callStaticLocalComponentMethodinit(methodname: string): Promise<void> {

  console.log("In callStaticLocalComponentMethodinit");

  try {
    let name = "Initializing";

    try {
      const dotnetloaded = await preloadDotNet();

      name = "something";

      if (dotnetloaded === true) {

        name = "Dotnet Loaded";
        // Call JSInvokable Function here ...
        name = await DotNet.invokeMethodAsync(
          "Blazor.Word.AddIn.Client",
          methodname,
          "Blazor Fan"
        );
      } else {
        name = "Init DotNet Failed";
      }
    } catch (error: unknown) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      name = errorMessage;
      console.error("Error during DotNet invocation: " + name);
    }

    await Word.run(async (context) => {
      const body = context.document.body;
      body.insertText(name, Word.InsertLocation.end);
      await context.sync();
    });

    console.log("Finished Initializing: " + name);
  } catch (error: unknown) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    console.error("Error in callStaticLocalComponentMethodinit:", errorMessage);
  } finally {
    console.log("Finish callStaticLocalComponentMethodinit");
  }
}

/**
 * Waits for the .NET runtime to be ready.
 * 
 * Uses a promise-based approach where the Blazor module signals readiness
 * via afterWebAssemblyStarted, eliminating the need for polling.
 * 
 * @param timeoutMs - Maximum time to wait for .NET to be ready (default: 10000ms)
 * @returns {Promise<boolean>} Returns true if the .NET runtime is ready, false if timeout.
 */
async function preloadDotNet(timeoutMs: number = 10000): Promise<boolean> {
  console.log("In preloadDotNet");

  try {
    const dotNetReadyPromise = (window as any).dotNetReady;

    if (!dotNetReadyPromise) {
      console.error("dotNetReady promise not found - Blazor module may not be loaded");
      return false;
    }

    // Race between the ready promise and a timeout
    const timeoutPromise = new Promise<never>((_, reject) =>
      setTimeout(() => reject(new Error("Timeout waiting for .NET runtime")), timeoutMs)
    );

    await Promise.race([dotNetReadyPromise, timeoutPromise]);
    console.log(".NET runtime is ready");
    return true;

  } catch (error: unknown) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    console.error("Error in preloadDotNet: " + errorMessage);
    return false;

  } finally {
    console.log("Finish preloadDotNet");
  }
}

/**
 * Calls the PrepareDocument method from ContentControls Blazor Page
 * @param event Office add-in command event
 * @returns A promise that resolves when the document preparation is complete
 */
async function callBlazorPrepareDocument(event: Office.AddinCommands.Event): Promise<void> {
  try {
    console.log("Running callBlazorPrepareDocument");

    // Ensure .NET runtime is loaded
    const dotnetloaded = await preloadDotNet();

    console.log("Running callBlazorPrepareDocument, preload: " + dotnetloaded);

    if (dotnetloaded === true) {
      // Call the JSInvokable PrepareDocument method
      await DotNet.invokeMethodAsync(
        "Blazor.Word.AddIn.Client",
        "PrepareDocument"
      );
      console.log("PrepareDocument completed successfully");
    } else {
      console.error("Init DotNet Failed");
    }
  } catch (error: unknown) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    console.error("Error calling PrepareDocument:", errorMessage);
  } finally {
    console.log("Finish callBlazorPrepareDocument");

    // Required: Let the platform know processing has completed
    if (event && typeof event.completed === 'function') {
      event.completed();
    }
  }
}

// Associate the functions with their named counterparts in the manifest XML.
Office.actions.associate("insertTextInWord", insertTextInWord);
Office.actions.associate("callBlazorOnHome", callBlazorOnHome);
Office.actions.associate("callBlazorOnCounter", callBlazorOnCounter);
Office.actions.associate("callBlazorPrepareDocument", callBlazorPrepareDocument);
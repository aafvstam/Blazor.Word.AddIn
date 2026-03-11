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
 * Writes the text from the Home Blazor Page to the PowerPoint slide.
 * Uses the "wasm" DotNetObjectReference (ClientCommandHandler) via WasmBridge.
 * @param {any} event
 */
async function callBlazorWasm(event: Office.AddinCommands.Event) {
  try {
    console.log("Running callBlazorWasm");
    await callDotNetMethod("wasm", "SayHelloWASM");
  }
  catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    console.error("Error in callBlazorWasm:", errorMessage);
  }
  finally {
    console.log("Finish callBlazorWasm");
  }
  // Be sure to indicate when the add-in command function is complete
  if (event && typeof event.completed === 'function') {
    event.completed();
  }
}
/**
 * Writes the text from the Counter Blazor Page to the PowerPoint slide.
 * Uses the "server" DotNetObjectReference (ServerCommandHandler) via ServerBridge.
 * @param {any} event
 */
async function callBlazorServer(event: Office.AddinCommands.Event) {
  try {
    console.log("Running callBlazorServer");
    await callDotNetMethod("server", "SayHelloServer");
  }
  catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    console.error("Error in callBlazorServer:", errorMessage);
  }
  finally {
    console.log("Finish callBlazorServer");
  }
  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
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
    console.log("Running callBlazorHome");
    await callDotNetMethod("wasm", "SayHelloHome");
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    console.error("Error in callBlazorHome:", errorMessage);
  } finally {
    console.log("Finish callBlazorHome");
  }

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
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
async function callBlazorPrepareDocument(event: Office.AddinCommands.Event): Promise<void> {

  try {
    console.log("Running callBlazorPrepareDocument");
    await callDotNetMethod("wasm", "PrepareDocument");
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    console.error("Error in callBlazorPrepareDocument:", errorMessage);
  } finally {
    console.log("Finish callBlazorPrepareDocument");
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
    await callDotNetMethod("server", "SayHelloCounter");
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
 * Invokes a .NET method on the named DotNetObjectReference and inserts the result
 * into a Word as a text.
 *
 * @param {string} bridgeName - The bridge name ("wasm" or "server") to look up in dotNetRefs.
 * @param {string} methodName - The name of the [JSInvokable] method to invoke on the handler.
 */
async function callDotNetMethod(bridgeName: string, methodName: string): Promise<void> {

  const t0 = performance.now();
  console.log(`In callDotNetMethod: bridge=${bridgeName}, method=${methodName}`);

  try {
    let name = "Initializing";

    try {
      const dotnetloaded = await preloadDotNet(bridgeName);

      if (dotnetloaded === true) {
        const dotNetRef = window.dotNetRefs.get(bridgeName);

        if (!dotNetRef) {
          name = `Bridge '${bridgeName}' not found in dotNetRefs`;
          console.error(name);
        } else {
          name = "Dotnet Loaded";
          // Call [JSInvokable] instance method on the DotNetObjectReference
          name = await dotNetRef.invokeMethodAsync<string>(methodName, "Blazor Fan");
        }
      } else {
        name = "Init DotNet Failed, methodName: " + methodName;
      }
    } catch (error: unknown) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      name = errorMessage;
      console.error("Error during DotNet invocation: " + name);
    }

    console.log(`callDotNetMethod: .NET call took ${(performance.now() - t0).toFixed(1)}ms, starting PowerPoint.run`);
    await insertText(name);
    console.log("Finished: " + name)
  } catch (error: unknown) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    console.error("Error in callDotNetMethod:", errorMessage);
  } finally {
    console.log("Finish callDotNetMethod");
  }
}

/**
 * Waits for a specific .NET bridge component to be ready.
 *
 * Each bridge (wasm, server) has its own promise that resolves when
 * that bridge's component signals readiness via signalDotNetReady(name, dotNetRef).
 *
 * @param bridgeName - The bridge to wait for ("wasm" or "server")
 * @param timeoutMs - Maximum time to wait for the bridge to be ready (default: 10000ms)
 * @returns {Promise<boolean>} Returns true if the bridge is ready, false if timeout.
 */
async function preloadDotNet(bridgeName: string, timeoutMs: number = 10000): Promise<boolean> {
  console.log(`In preloadDotNet: waiting for ${bridgeName} bridge`);

  try {
    const bridgePromise = window.dotNetReady?.[bridgeName];

    if (!bridgePromise) {
      console.error(`dotNetReady.${bridgeName} promise not found - Blazor module may not be loaded`);
      return false;
    }

    // Race between the bridge ready promise and a timeout
    await Promise.race([
      bridgePromise,
      new Promise<never>((_, reject) => {
        AbortSignal.timeout(timeoutMs).addEventListener("abort", () =>
          reject(new Error(`Timeout waiting for ${bridgeName} bridge`))
        );
      }),
    ]);
    console.log(`${bridgeName} bridge is ready`);
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
 * Inserts text into a text box on the currently selected PowerPoint slide.
 *
 * @param text - The text to insert into the slide
 * @param options - Optional positioning options for the text box
 */
async function insertText(text: string) {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.insertText(text, Word.InsertLocation.end);
    await context.sync();
  });
}

// Associate the functions with their named counterparts in the manifest XML.
Office.actions.associate("insertTextInWord", insertTextInWord);
Office.actions.associate("callBlazorOnHome", callBlazorOnHome);
Office.actions.associate("callBlazorWasm", callBlazorWasm);
Office.actions.associate("callBlazorServer", callBlazorServer);
Office.actions.associate("callBlazorPrepareDocument", callBlazorPrepareDocument);
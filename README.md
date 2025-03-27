---
page_type: sample
urlFragment: word-hybrid-blazor-add-in
products:
  - office
  - office-word
languages:
  - javascript
  - C#
extensions:
  contentType: samples
  technologies: 
    - Add-ins
  createdDate: '02/21/2025 10:00:00 PM'
description: 'Create a Hybrid Blazor Word add-in showcasing some samples.'
---

# Create a Hybrid Blazor Word add-in

This sample shows how to build a Word add-in using .NET Blazor Hybrid technologies. Blazor Webassembly allows you to build Office Add-ins using .NET, C#, and JavaScript to interact with the Office JS API. The add-in uses JavaScript to work with the document and Office JS APIs, but you build the user interface and all other non-Office interactions in C# and .NET Core Blazor technologies.

- Work with Hybrid Blazor Webassembly in the context of Office.
- Build cross-platform Office Add-ins using Hybrid Blazor, C# and JavaScript Interop.
- Initialize the Office JavaScript API library in Hybrid Blazor context.
- Interact with Word to manipulate Documents.
- Interact with document content through Office JavaScript APIs.
- Interact with methods defined on the Hybrid Blazor Pages.
- Interop between OfficeJS - JavaScript - C# and back to JavaScript.

## Applies to

- Word on the web, Windows, and Mac.

## Prerequisites

- Microsoft 365 - Get a [free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) that provides a renewable 90-day Microsoft 365 E5 developer subscription.

## Run the sample

1. Download or clone the [Office Add-ins samples repository](https://github.com/OfficeDev/Office-Add-in-samples).
1. Open Visual Studio 2022 and open the: **Office-Add-in-samples\Samples\blazor-add-in\Blazor.Word.AddIn\Blazor.Word.AddIn.sln** solution.
1. Choose **Debug** > **Start Debugging**. Or press <kbd>F5</kbd> to start the solution and sideload the Add-in and start Word.
1. [Optionally] You can use the Terminal to sideload the Add-in and start Word using the command **npm run start-local** separately
1. When Word opens, choose **Sample Add-in** > **Show task pane** (if not already open).
1. Try out the controls on the task panes.
1. Try out the custom buttons on the **Sample Add-in** tab on the ribbon.

## Understand an Office Add-in in a Hybrid Blazor Context
An Office Add-in is a web application that extends Office with additional functionality for the user. For example, an add-in can add ribbon buttons, a task pane, or a content pane with the functionality you want. Because an Office Add-in is a web application, you must provide a web server to host the files.  

Building the Office Add-in as a Hybrid Blazor Application it allows you to build a .NET Core compliant website that interacts with the Office JS APIs using WebAssembly. If your background is with VBA, VSTO, or COM add-in development, you may find that building Office Add-ins using Hybrid Blazor is a familiar development technique.  

By using Hybrid Blazor, you can select which parts of your add-in you want to build in C# and which parts you want to build in JavaScript. This allows you to leverage the strengths of both languages and build a more robust add-in that can run parts that need a higher security level at the server and only bring the parts that interact with Office to the client.  

This sample also implemented the Blazor Fluent UI components to show how you can use the Fluent UI components in your add-in. The Blazor Fluent UI components are a set of UI components that implement the Fluent Design System. The Fluent UI components contains options to easily implement Themes to your Add-in. To switch between themes select the Theme option in the Ribbon, or the Theme option in the task pane menu.   

This sample also implemented TypeScript as a language to interact with the Office JS API. TypeScript is a superset of JavaScript that compiles to plain JavaScript. TypeScript is designed for the development of large applications and can be used to develop Office Add-ins. TypeScript is a language that is easy to learn and can be used to build Office Add-ins that interact with the Office JS API. TypeScript allows you to develop strong typed code that can detect compile issues before the code is deployed where often JavaScript issues only are discovered at runtime in production.  

## Key parts of this sample
This sample runs a Hybrid Blazor Application that  runs cross-platform in various browsers supporting WASM (Webassembly). The part that is using WebAssembly demonstrates some basic Word functions using the Document.

The purpose of this sample is to show you how to build and interact with the Blazor, C# and JavaScript Interop options. If you're looking for more examples of interacting with Word and Office JS APIs, see [Script Lab](https://aka.ms/getscriptlab).

### Blazor pages
The **Pages** folder contains the Blazor pages, and is based on the generic Blazor demo application provided in Visual Studio. It contains similar pages such as Home, Counter and Weather. The Home page is implemented as **Home.razor**. As pages can run both on the Server side or the Client side they can be defined in either the Add-in project or the Client project.  

Each **.razor** page can contain code-behind pages, for example, named **Home.razor.cs** and **Home.razor.ts**. The C# file first establishes an interop connection with the JavaScript file (generated by TypeScript).

```csharp
protected override async Task OnAfterRenderAsync(bool firstRender)
{
  if (firstRender)
  {
    JSModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./Pages/Home.razor.js");
  }
}
```

##### TODO - Implement TypeScript Interop Functions
For any events that need to interact with the Office document, the C# file calls through interop to the JavaScript (generated by TypeScript) file.

```csharp
private async Task InsertParagraph() =>
  await JSModule.InvokeVoidAsync("insertParagraph");
```

The JavaScript runs the code to interact with the document and returns.

```javascript
export function insertParagraph() {
  return Word.run((context) => {
    // insert a paragraph at the start of the document.
    const paragraph = context.document.body.insertParagraph(
      'Hello World from Blazor',
      Word.InsertLocation.start
    );

    // sync the context to run the previous API call, and return.
    return context.sync();
  });
}
```

The fundamental pattern includes the following steps.

1. Call **JSRuntime.InvokeAsync** to set up the interop between C# and JavaScript.
1. Use **JSModule.InvokeVoidAsync** to call JavaScript functions from your C# code.
1. Call Office JS APIs to interact with the document from JavaScript code.

### Blazor interop with Add-in Commands
This sample shows how to use Blazor with custom buttons on the ribbon. The buttons call the same functions that are defined on the task pane. This sample is configured to use the shared runtime which is required for this interop to work correctly.

## Debugging
##### TODO - Add more content here
This sample is configured to support debugging both JavaScript and C# files. New Blazor projects need the following file updates to support C# debugging.

1. In the **launchSettings.json** file of the web project, make sure all instances of `launchBrowser` are set to `false`.
1. In the **<projectName>.csproj.user** file of the add-in project, add the `<BlazorAppUrl>` and `<InspectUri>` elements as shown in the following example XML.

**Note:** The port number in the following XML is 7215. You must change it to the port number specified in the **launchSettings.json** file for your web project.

```xml
<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="Current" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <PropertyGroup>
    <BlazorAppUrl>https://localhost:7215/</BlazorAppUrl>
    <InspectUri>{wsProtocol}://{url.hostname}:{url.port}/_framework/debug/ws-proxy?browser={browserInspectUri}</InspectUri>
  </PropertyGroup>
</Project>
```

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Solution

| Solution                                 | Authors                                                                 |
| ---------------------------------------- | ----------------------------------------------------------------------- |
| Create a Blazor Hybrid Word add-in | [Maarten van Stam](https://mvp.microsoft.com/en-us/PublicProfile/33535) |

## Version history

| Version | Date                | Comments           |
| ------- | ------------------- | ------------------ |
| 1.0     | XXXXXXXX XX, 2025   | Work In Progress   |

## Extra Credits
Additional contributions to this sample were made by [Rudy Cortembert](https://github.com/RudyCo/MyAspireBlazorAddIn).  
Rudy adopted our earlier Blazor Add-in sample and helped to made it work with the new Blazor Hybrid model.

## Copyright

Copyright(c) Maarten van Stam. All rights reserved.Licensed under the MIT License.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

**Note**: The index.html file contains an image URL that tracks diagnostic data for this sample add-in. Please remove the image tag if you reuse this sample in your own code project.

**TO DO: Update the image URL below with the correct telemetry URL.**
<!-- 
    <img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/blazor-add-in/word-blazor-hybrid-add-in" /> 
-->

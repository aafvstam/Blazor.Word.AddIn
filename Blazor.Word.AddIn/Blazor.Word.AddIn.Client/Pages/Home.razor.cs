/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;

using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace Blazor.Word.AddIn.Client.Pages;

/// <summary>
/// Starter class to demo how to insert a paragraph
/// </summary>
[SupportedOSPlatform("browser")]
public partial class Home : ComponentBase
{
    private bool HostInformation;

    [JSImport("IsRunningInHost", "Home")]
    internal static partial Task<bool> OfficeOnReady();

    protected override async Task OnAfterRenderAsync(bool firstRender)
    {
        if (firstRender)
        {
            try
            {
                await JSHost.ImportAsync("Home", "../Pages/Home.razor.js");
                Console.WriteLine($"Imported Home module");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error importing Home module: {ex.Message}");
            }
            
            HostInformation = await OfficeOnReady();
            Console.WriteLine($"Home HostInformation: {HostInformation}");

            if (HostInformation)
            {
                StateHasChanged();
            }
        }
    }

    [JSImport("insertParagraph", "Home")]
    internal static partial Task InsertParagraph();

    [JSInvokable]
    public static Task<string> SayHelloHome(string name)
    {
        Console.WriteLine("Invoking SayHelloHome");
        return Task.FromResult($"Hello Home, {name} from Home Page!");
    }

    [JSInvokable]
    public static Task<string> PreloaderDummy()
    {
        Console.WriteLine("Invoking PreloaderDummy");
        return Task.FromResult("Loaded");
    }
}
/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
using Blazor.Word.AddIn.Client.Model;

using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;

namespace Blazor.Word.AddIn.Client.Pages;

public partial class Home : ComponentBase
{
    private HostInformation hostInformation = new HostInformation();

    [Inject, AllowNull]
    private IJSRuntime JSRuntime { get; set; }
    private IJSObjectReference JSModule { get; set; } = default!;

    protected override async Task OnAfterRenderAsync(bool firstRender)
    {
        if (firstRender)
        {
            hostInformation = await JSRuntime.InvokeAsync<HostInformation>("Office.onReady");

            Debug.WriteLine("Hit OnAfterRenderAsync in Home.razor.cs!");
            Console.WriteLine("Hit OnAfterRenderAsync in Home.razor.cs in Console!");
            JSModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./Pages/Home.razor.js");

            if (hostInformation.IsInitialized)
            {
                StateHasChanged();
            }
        }
    }

    /// <summary>
    /// Function to create a new slide in the Word presentation.
    /// </summary>
    private async Task CreateSlide() =>
        await JSModule.InvokeVoidAsync("createSlide");


    [JSInvokable]
    public static Task<string> SayHelloHome(string name)
    {
        return Task.FromResult($"Hello Home, {name} from Home Page!");
    }

    [JSInvokable]
    public static Task<string> PreloaderDummy()
    {
        return Task.FromResult("Loaded");
    }
}
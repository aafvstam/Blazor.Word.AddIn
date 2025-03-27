﻿namespace Blazor.Word.AddIn.Client.Model
{
    public class HostInformation
    {
        public string? Host { get; set; }
        public string? Platform { get; set; }
        public bool IsInitialized => Host != null && Platform != null;
    }
}

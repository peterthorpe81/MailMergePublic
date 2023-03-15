using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using MailMerge;
using MudBlazor.Services;

var builder = WebAssemblyHostBuilder.CreateDefault(args);
builder.RootComponents.Add<App>("#app");
builder.RootComponents.Add<HeadOutlet>("head::after");

builder.Services.AddScoped(sp => new HttpClient { BaseAddress = new Uri(builder.HostEnvironment.BaseAddress) });

builder.Services.AddMudServices();

//settings as named file
var settings = new ClientAppSettings();
builder.Configuration.Bind("AzureAd:ProviderOptions:Authentication", settings);
builder.Services.AddSingleton(settings);

builder.Services.AddMsalAuthentication(options =>
{
    builder.Configuration.Bind("AzureAd", options);
    //options.ProviderOptions.DefaultAccessTokenScopes.Add("Mail.Send");
    //options.ProviderOptions.AdditionalScopesToConsent.Add("https://graph.microsoft.com/Files.Read");
});

var baseUrl = string.Join("/", 
    builder.Configuration.GetSection("MicrosoftGraph")["BaseUrl"], 
    builder.Configuration.GetSection("MicrosoftGraph")["Version"]);
var scopes = builder.Configuration.GetSection("MicrosoftGraph:Scopes")
    .Get<List<string>>();

builder.Services.AddGraphClient(baseUrl, scopes);


await builder.Build().RunAsync();

using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Graph;
using Azure.Identity;
using System;

var host = new HostBuilder()
    .ConfigureFunctionsWebApplication()
    .ConfigureServices(services =>
    {
        // Spustí monitorování funkcí (Application Insights)
        services.AddApplicationInsightsTelemetryWorkerService();
        services.ConfigureFunctionsApplicationInsights();

        // Zde zaregistrujeme GraphServiceClient jako Singleton.
        // To zajistí, že se připojení vytvoří jen jednou při startu (Warm Start)
        // a nebude se zbytečně otevírat znovu při každém přiložení karty.
        services.AddSingleton<GraphServiceClient>(sp =>
        {
            var tenantId = Environment.GetEnvironmentVariable("AzureAd_TenantId");
            var clientId = Environment.GetEnvironmentVariable("AzureAd_ClientId");
            var clientSecret = Environment.GetEnvironmentVariable("AzureAd_ClientSecret");

            // Kontrola pro lokální vývoj, aby aplikace nespadla hned při startu, pokud chybí klíče
            if (string.IsNullOrEmpty(tenantId) || string.IsNullOrEmpty(clientId) || string.IsNullOrEmpty(clientSecret))
            {
                // Pokud klíče chybí, vrátíme prázdného klienta (prozatím)
                return new GraphServiceClient(new ClientSecretCredential("placeholder", "placeholder", "placeholder"));
            }

            var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
            return new GraphServiceClient(credential, new[] { "https://graph.microsoft.com/.default" });
        });
    })
    .Build();

host.Run();
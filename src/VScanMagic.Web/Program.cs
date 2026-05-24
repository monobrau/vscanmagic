using VScanMagic.ConnectSecure;
using VScanMagic.Core;
using VScanMagic.Core.Services;
using VScanMagic.Data;
using VScanMagic.Reports;
using VScanMagic.Review;
using VScanMagic.Review.Models;
using VScanMagic.Review.Services;
using VScanMagic.Review.Storage;
using VScanMagic.Web.Api;
using VScanMagic.Web.Components;
using VScanMagic.Web.Services;
using Serilog;

Log.Logger = new LoggerConfiguration()
    .WriteTo.Console()
    .CreateLogger();

var builder = WebApplication.CreateBuilder(args);
builder.Host.UseSerilog();

builder.Services.AddVScanMagicCore();
builder.Services.AddVScanMagicData();
builder.Services.AddVScanMagicReview();
builder.Services.AddVScanMagicReports();
builder.Services.AddVScanMagicConnectSecure();
builder.Services.AddSingleton<VulnerabilityPipeline>();
builder.Services.AddSingleton<ExportOrchestrator>();
builder.Services.AddSingleton<AppRestartService>();

builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

builder.Services.AddControllers();

var app = builder.Build();

var savedCreds = app.Services.GetRequiredService<SettingsService>().LoadConnectSecureCredentials();
if (!string.IsNullOrWhiteSpace(savedCreds.BaseUrl))
    app.Services.GetRequiredService<ConnectSecureClient>().Configure(savedCreds);

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseAntiforgery();

app.MapControllers();
app.MapReviewSessionApi();
app.MapLegacyReportApi();
app.MapConnectSecureDownloadApi();
app.MapAdminApi();
app.MapPatchApi();

app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

var bind = Environment.GetEnvironmentVariable("VSCANMAGIC_API_BIND") ?? "127.0.0.1";
var port = Environment.GetEnvironmentVariable("VSCANMAGIC_PORT") ?? "8080";
app.Urls.Add($"http://{bind}:{port}");

Log.Information("VScanMagic Web starting on http://{Bind}:{Port}", bind, port);
app.Run();

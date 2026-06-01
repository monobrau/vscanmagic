using VScanMagic.ConnectSecure;
using VScanMagic.ConnectWiseManage;
using VScanMagic.Core;
using VScanMagic.Core.Services;
using VScanMagic.Data;
using VScanMagic.Reports;
using VScanMagic.Review;
using VScanMagic.Review.Models;
using VScanMagic.Review.Services;
using VScanMagic.Review.Storage;
using Microsoft.AspNetCore.Components.Server;
using VScanMagic.Web.Api;
using VScanMagic.Web.Components;
using VScanMagic.Web.Services;
using Serilog;

Log.Logger = new LoggerConfiguration()
    .WriteTo.Console()
    .CreateLogger();

var bind = Environment.GetEnvironmentVariable("VSCANMAGIC_API_BIND") ?? "127.0.0.1";
var port = Environment.GetEnvironmentVariable("VSCANMAGIC_PORT") ?? "8080";
var loopbackOnly = AppRestartSupport.IsLocalBind(bind);

var builder = WebApplication.CreateBuilder(args);
builder.Host.UseSerilog();
Log.Information("Timestamps use local timezone: {TimeZone} (UTC{Offset})",
    TimeZoneInfo.Local.DisplayName,
    TimeZoneInfo.Local.GetUtcOffset(DateTime.Now).ToString(@"hh\:mm"));

builder.Services.AddVScanMagicCore();
builder.Services.AddVScanMagicData();
builder.Services.AddVScanMagicReview();
builder.Services.AddVScanMagicReports();
builder.Services.AddVScanMagicConnectSecure();
builder.Services.AddVScanMagicConnectWiseManage();
builder.Services.AddSingleton<ExportOrchestrator>();
builder.Services.AddSingleton<AppRestartService>();
builder.Services.AddSingleton<NativeFolderPickerService>();
builder.Services.AddSingleton<OutlookDeliverableDraftService>();
builder.Services.AddSingleton<BulkReviewJobService>();
builder.Services.AddSingleton<SessionSupplementalReportService>();
builder.Services.AddScoped<CompanyListService>();
builder.Services.AddScoped<LoadTimingDisplay>();

builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

builder.Services.Configure<CircuitOptions>(options =>
{
    options.DetailedErrors = builder.Environment.IsDevelopment()
        || string.Equals(Environment.GetEnvironmentVariable("VSCANMAGIC_DETAILED_ERRORS"), "1", StringComparison.OrdinalIgnoreCase);
    // Long report downloads (All Vulnerabilities) can take several minutes; keep circuit alive over LAN/VPN.
    options.DisconnectedCircuitRetentionPeriod = TimeSpan.FromMinutes(15);
    options.JSInteropDefaultCallTimeout = TimeSpan.FromMinutes(10);
});

builder.Services.AddSignalR(options =>
{
    options.ClientTimeoutInterval = TimeSpan.FromMinutes(5);
    options.KeepAliveInterval = TimeSpan.FromSeconds(15);
});

builder.Services.AddAntiforgery(options =>
{
    options.Cookie.SecurePolicy = CookieSecurePolicy.SameAsRequest;
});

builder.Services.AddControllers();

var app = builder.Build();

var savedCreds = app.Services.GetRequiredService<SettingsService>().LoadConnectSecureCredentials();
if (!string.IsNullOrWhiteSpace(savedCreds.BaseUrl))
    app.Services.GetRequiredService<ConnectSecureClient>().Configure(savedCreds);

var manageStore = app.Services.GetRequiredService<ConnectWiseManageSettingsStore>();
var manageCreds = manageStore.LoadCredentials();
if (!string.IsNullOrWhiteSpace(manageCreds.ApiUrl) &&
    !string.IsNullOrWhiteSpace(manageCreds.CompanyId) &&
    !string.IsNullOrWhiteSpace(manageCreds.PublicKey) &&
    !string.IsNullOrWhiteSpace(manageCreds.PrivateKey) &&
    !string.IsNullOrWhiteSpace(manageCreds.ClientId))
{
    try
    {
        app.Services.GetRequiredService<ConnectWiseManageClient>().Configure(manageCreds);
    }
    catch (Exception ex)
    {
        Log.Warning(ex, "Saved ConnectWise Manage credentials could not be applied at startup.");
    }
}

await app.Services.GetRequiredService<BulkReviewJobService>().RecoverInterruptedJobsAsync();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    if (loopbackOnly)
        app.UseHsts();
}

if (loopbackOnly)
    app.UseHttpsRedirection();
else
    Log.Warning("LAN bind ({Bind}): HTTPS redirection disabled. Expose only on trusted networks.", bind);

app.UseWebSockets();
app.UseStaticFiles();
app.UseAntiforgery();

app.MapControllers();
app.MapReviewSessionApi();
app.MapLegacyReportApi();
app.MapConnectSecureDownloadApi();
app.MapAdminApi();
app.MapPatchApi();
app.MapConnectWiseManageApi();

app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Urls.Add($"http://{bind}:{port}");

Log.Information("VScanMagic Web starting on http://{Bind}:{Port}", bind, port);
app.Run();

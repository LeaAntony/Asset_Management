using Asset_Management.Function;
using Asset_Management.Service;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.IdentityModel.Tokens;
using System.Text;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authentication.OAuth;
using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.AspNetCore.Authentication;
using System.Security.Claims;
using System.Net.Http.Headers;
using System.Text.Json;
using Asset_Management.Models;
using Microsoft.AspNetCore.DataProtection;
using Microsoft.AspNetCore.WebUtilities;
using System.Web;

var builder = WebApplication.CreateBuilder(args);

var env = builder.Environment;

builder.Configuration
    .SetBasePath(Directory.GetCurrentDirectory())
    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
    .AddJsonFile($"appsettings.{builder.Environment.EnvironmentName}.json", optional: true, reloadOnChange: true)
    .AddEnvironmentVariables();

var configuration = builder.Configuration;

string connectionString = "Data Source=localhost;Initial Catalog=Asset_Management;Integrated Security=true;Persist Security Info=True;MultipleActiveResultSets=true";
builder.Services.AddDbContext<ApplicationDbContext>(options =>
        options.UseSqlServer(
                    connectionString,
                    b => b.MigrationsAssembly(typeof(ApplicationDbContext).Assembly.FullName)));
builder.Services.AddControllersWithViews();
builder.Services.AddTransient<ITokenService, TokenService>();

builder.Services.AddAuthentication(options =>
{
    options.DefaultScheme = CookieAuthenticationDefaults.AuthenticationScheme;
})
.AddCookie(
    options =>
    {
        options.Cookie.Name = "ping";
        options.Cookie.Path = "/";
        options.Cookie.SecurePolicy = CookieSecurePolicy.Always;
        options.Cookie.HttpOnly = true;
        options.SlidingExpiration = true;
        options.Events = new CookieAuthenticationEvents
        {
            OnRedirectToAccessDenied = context =>
            {
                var pathBase = context.HttpContext.Request.PathBase;
                context.Response.Redirect(pathBase + "/Home/Login");
                return Task.CompletedTask;
            },
            OnRedirectToLogin = context =>
            {
                var pathBase = context.HttpContext.Request.PathBase;
                context.Response.Redirect(pathBase + "/Home/Login");
                return Task.CompletedTask;
            }
        };
    }
);


builder.Services.AddAuthorization(options =>
{
    options.AddPolicy("RequireRequestor", policy => policy.RequireClaim("asset_management_level", "requestor"));
    options.AddPolicy("RequireApprover", policy => policy.RequireClaim("asset_management_level", "approver"));
    options.AddPolicy("RequireFinance", policy => policy.RequireClaim("asset_management_level", "finance"));
    options.AddPolicy("RequireSecurity", policy => policy.RequireClaim("asset_management_level", "security"));
    options.AddPolicy("RequireAdmin", policy => policy.RequireClaim("asset_management_level", "admin"));
    options.AddPolicy("RequireRequestorApprover", policy => policy.RequireAssertion(context => context.User.HasClaim("asset_management_level", "requestor") || context.User.HasClaim("asset_management_level", "approver")));
    options.AddPolicy("RequireFinanceApprover", policy => policy.RequireAssertion(context => context.User.HasClaim("asset_management_level", "finance") || context.User.HasClaim("asset_management_level", "approver")));
    options.AddPolicy("RequireFinanceAdmin", policy => policy.RequireAssertion(context =>
                    context.User.HasClaim("asset_management_level", "finance") ||
                    context.User.HasClaim("asset_management_level", "admin") ||
                    context.User.HasClaim("asset_management_role_manage_user", "1")));
    options.AddPolicy("RequireAny", policy => policy.RequireAssertion(context =>
                    context.User.HasClaim("asset_management_level", "requestor") ||
                    context.User.HasClaim("asset_management_level", "approver") ||
                    context.User.HasClaim("asset_management_level", "security") ||
                    context.User.HasClaim("asset_management_level", "admin") ||
                    context.User.HasClaim("asset_management_level", "finance")));
});

builder.Services.AddMvc();
builder.Services.AddSession();
builder.Services.AddHttpContextAccessor();
builder.Services.AddHttpClient();
builder.Services.AddSingleton<ImportExportFactory>();

var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

app.UseAuthentication();
app.UseAuthorization();
app.UseSession();

app.MapControllerRoute(
    name: "callback",
    pattern: "callback",
    defaults: new { controller = "Callback", action = "HandleCallback" });

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.Run();

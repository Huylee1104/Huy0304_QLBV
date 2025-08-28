using C0304.Db.Models;
using Microsoft.EntityFrameworkCore;
using QuestPDF.Infrastructure;
using S0304BangKeThu.Services;
using S0304HTTT.Services;
using S0304NhanVien.Services;
using S0304ThongTinDoanhNghiep.Services;
using S0304BBCHoaDonDienTuDV.Services;

var builder = WebApplication.CreateBuilder(args);

// -------------------- KHAI BÁO DB CONTEXT --------------------
builder.Services.AddDbContext<M0304Context>(options =>
options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection")));

builder.Services.AddScoped<I0304ThongTinDoanhNghiep, S0304ThongTinDoanhNghiepService>();
builder.Services.AddScoped<I0304NhanVienService, S0304NhanVienService>();
builder.Services.AddScoped<I0304HTTTService, S0304HTTTService>();
builder.Services.AddScoped<I0304BangKeThuService, S0304BangKeThuService>();
builder.Services.AddScoped<I0304BBCHoaDonDienTuDVService, S0304BBCHoaDonDienTuService>();

builder.Services.AddDistributedMemoryCache(); // Bộ nhớ tạm cho session
builder.Services.AddSession(options =>
{
    options.IdleTimeout = TimeSpan.FromMinutes(30); // Thời gian timeout
    options.Cookie.HttpOnly = true;
    options.Cookie.IsEssential = true;
});

// Cấu hình logging
builder.Logging.ClearProviders();
builder.Logging.AddConsole(); // hoặc AddDebug nếu muốn log ra Output window
builder.Logging.AddDebug();
builder.Services.AddControllersWithViews();

// Add services to the container.
builder.Services.AddControllersWithViews();
builder.Services.AddHttpContextAccessor();

var app = builder.Build();
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

QuestPDF.Settings.License = LicenseType.Community;

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();
app.UseSession();
app.UseAuthorization();

app.UseEndpoints(endpoints =>
{
    endpoints.MapControllers(); // Cho phép attribute routing hoạt động
    endpoints.MapControllerRoute(
        name: "default",
        pattern: "{controller=Home}/{action=Index}/{id?}");
});

app.Run();

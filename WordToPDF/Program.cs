using Microsoft.EntityFrameworkCore;
using WordToPDF.Data;
using WordToPDF.Services;

var builder = WebApplication.CreateBuilder(args);


builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();


builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowAll", policy =>
    {
        policy.AllowAnyOrigin()
              .AllowAnyHeader()
              .AllowAnyMethod();
    });
});


builder.Services.AddDbContext<ApplicationDbContext>(options =>
    options.UseOracle(builder.Configuration.GetConnectionString("DefaultConnection")));

// SQL Server context for the Employee/Education/WordTemplate (PDF converter) flow
builder.Services.AddDbContext<EmployeeDbContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("SqlServerConnection")));


builder.Services.AddScoped<IPDFService, PDFService>();
builder.Services.AddScoped<IAppointmentLetterService, AppointmentLetterService>();
builder.Services.AddScoped<ITemplateService, SqlTemplateService>();
builder.Services.AddScoped<IPaySlipService, PaySlipService>();
builder.Services.AddScoped<IWaterBillService, WaterBillService>();

var app = builder.Build();

app.UseSwagger();
app.UseSwaggerUI(c =>
{
    c.SwaggerEndpoint("/swagger/v1/swagger.json", "My API V1");
    c.RoutePrefix = "swagger";
});


app.UseCors("AllowAll");

//app.UseHttpsRedirection();
app.UseAuthorization();

app.MapControllers();

app.Run();

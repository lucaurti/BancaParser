using BancaParser.Host.Helpers;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.IdentityModel.Tokens;
using Microsoft.OpenApi.Models;
using System.Text;

namespace BancaParser.Host
{

  public class Program
  {
    public static void Main(string[] args)
    {
      var builder = WebApplication.CreateBuilder(args);
      string path = Path.Combine(BancaParser.Core.Utility.PATH_SHARED_FOLDER, "dati", "banca-parser", "appsettings.json");
      FileInfo fileInfo = new FileInfo(path);
      if (File.Exists(path))
      {
        builder.Configuration
          .AddJsonFile(fileInfo.FullName, optional: true, reloadOnChange: true);
      }
      else
      {
        builder.Configuration
          .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
      }
      builder.Services.Configure<Config>(builder.Configuration.GetSection("AppSettings"));
      var appSettingsSection = builder.Configuration.GetSection("AppSettings");
      var secretKey = appSettingsSection["Secret"];

      // ===== AUTENTICAZIONE JWT =====
      builder.Services.AddAuthentication(options =>
    {
      options.DefaultAuthenticateScheme = JwtBearerDefaults.AuthenticationScheme;
      options.DefaultChallengeScheme = JwtBearerDefaults.AuthenticationScheme;
    })
    .AddJwtBearer(options =>
     {
       options.TokenValidationParameters = new TokenValidationParameters
       {
         ValidateIssuer = true,
         ValidateAudience = true,
         ValidateLifetime = true,
         ValidateIssuerSigningKey = true,
         ValidIssuer = "banca-parser",
         ValidAudience = "banca-parser-users",
         IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(secretKey))
       };
     });

      // ===== CONTROLLERS =====
      builder.Services.AddControllers();

      // ===== SWAGGER =====
      builder.Services.AddEndpointsApiExplorer();
      builder.Services.AddSwaggerGen(c =>
      {
        c.SwaggerDoc("v1", new OpenApiInfo
        {
          Title = "My Secure API (.NET 9)",
          Version = "v1"
        });

        // Configurazione schema JWT Bearer per Swagger
        var jwtScheme = new OpenApiSecurityScheme
        {
          Name = "Authorization",
          Description = "Inserisci il token JWT (es: Bearer eyJhbGci...)",
          In = ParameterLocation.Header,
          Type = SecuritySchemeType.Http,
          Scheme = "bearer",
          BearerFormat = "JWT",
          Reference = new OpenApiReference
          {
            Type = ReferenceType.SecurityScheme,
            Id = JwtBearerDefaults.AuthenticationScheme
          }
        };

        c.AddSecurityDefinition(JwtBearerDefaults.AuthenticationScheme, jwtScheme);
        c.AddSecurityRequirement(new OpenApiSecurityRequirement
          {
        { jwtScheme, new string[] { } }
          });
      });

      var app = builder.Build();

      // ===== MIDDLEWARE PIPELINE =====

      app.UseDeveloperExceptionPage();
      app.UseSwagger();
      app.UseSwaggerUI();


      app.UseHttpsRedirection();

      app.UseAuthentication();
      app.UseAuthorization();

      app.MapControllers();

      app.Run();
    }
  }
}

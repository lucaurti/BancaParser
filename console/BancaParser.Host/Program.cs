using Microsoft.IdentityModel.Tokens;
using Microsoft.OpenApi.Models;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text;

namespace BancaParser.Host
{
  public class Program
  {
    public static void Main(string[] args)
    {
      var builder = WebApplication.CreateBuilder(args);

      // Add services to the container.
      builder.Services.AddEndpointsApiExplorer();
      builder.Services.AddSwaggerGen(options =>
      {
        options.SwaggerDoc("v1", new OpenApiInfo
        {
          Title = "Banca Parser Api",
          Version = "v1",
          Description = "Banca Parser Api",
        });
        options.AddSecurityDefinition("Bearer", new Microsoft.OpenApi.Models.OpenApiSecurityScheme
        {
          Name = "Authorization",
          Type = Microsoft.OpenApi.Models.SecuritySchemeType.Http,
          Scheme = "Bearer",
          BearerFormat = "JWT",
          In = Microsoft.OpenApi.Models.ParameterLocation.Header,
          Description = "JWT Authorization header using the Bearer scheme."
        });
        options.AddSecurityRequirement(new Microsoft.OpenApi.Models.OpenApiSecurityRequirement
        {
          {
            new Microsoft.OpenApi.Models.OpenApiSecurityScheme
            {
              Reference = new Microsoft.OpenApi.Models.OpenApiReference
              {
                Type = Microsoft.OpenApi.Models.ReferenceType.SecurityScheme,
                Id = "Bearer"
              }
            },
            new string[] {}
          }
        });
      });
      builder.Services.AddControllers();

      var app = builder.Build();

      // Attiva Swagger solo in sviluppo (opzionale)
      if (app.Environment.IsDevelopment())
      {
        app.UseSwagger();
        app.UseSwaggerUI(c =>
        {
          c.SwaggerEndpoint("/swagger/v1/swagger.json", "Banca Parser Api v1");
          c.RoutePrefix = string.Empty; // Mostra Swagger UI sulla root "/"
        });
      }


      // Configure the HTTP request pipeline.

      app.UseAuthorization();
      app.UseMiddleware<JwtMiddleware>();

      app.MapControllers();

      app.Run();
    }
  }

  public class JwtMiddleware
  {
    private readonly RequestDelegate _next;

    public JwtMiddleware(RequestDelegate next)
    {
      _next = next;
    }

    public async Task Invoke(HttpContext context)
    {
      var token = context.Request.Headers["Authorization"].FirstOrDefault()?.Split(" ").Last();

      if (token != null)
      {
        attachUserToContext(context, token);
      }
      await _next(context);
    }

    public void attachUserToContext(HttpContext context, string token)
    {
      try
      {
        JwtSecurityToken jwtToken = GetJwtSecurityToken(token);
        string userId = jwtToken.Claims.First(x => x.Type == "id").Value.ToString();
        // attach user to context on successful jwt validation
        //var user = accountService.GetUserById(userId);
        //if (user == null)
        //  throw new Exception("User not found");
        //context.Items["User"] = user;
      }
      catch (Exception)
      {
        context.Items["User"] = null;
        context.Items["JwtValidationError"] = true;
      }
    }

    public static string generateJwtToken(string userId, DateTime expiresTime)
    {
      // generate token that is valid for 1 days
      var tokenHandler = new JwtSecurityTokenHandler();
      var key = Encoding.ASCII.GetBytes("SECRET_KET_MD5_BY_BANCA_PARSER_E4395424-987F-45A1-A3B3-79CF03134C8D_LUCLASH!");
      var tokenDescriptor = new SecurityTokenDescriptor
      {
        Subject = new ClaimsIdentity(new[] { new Claim("id", userId) }),
        Expires = expiresTime,
        SigningCredentials = new SigningCredentials(new SymmetricSecurityKey(key), SecurityAlgorithms.HmacSha512Signature)
      };
      var token = tokenHandler.CreateToken(tokenDescriptor);
      return tokenHandler.WriteToken(token);
    }

    public static JwtSecurityToken GetJwtSecurityToken(string token)
    {
      var tokenHandler = new JwtSecurityTokenHandler();
      var key = Encoding.ASCII.GetBytes("SECRET_KET_MD5_BY_BANCA_PARSER_E4395424-987F-45A1-A3B3-79CF03134C8D_LUCLASH!");
      tokenHandler.ValidateToken(token, new TokenValidationParameters
      {
        ValidateIssuerSigningKey = true,
        IssuerSigningKey = new SymmetricSecurityKey(key),
        ValidateIssuer = false,
        ValidateAudience = false,
        // set clockskew to zero so tokens expire exactly at token expiration time (instead of 5 minutes later)
        ClockSkew = TimeSpan.Zero
      }, out SecurityToken validatedToken);

      return (JwtSecurityToken)validatedToken;
    }
  }
}

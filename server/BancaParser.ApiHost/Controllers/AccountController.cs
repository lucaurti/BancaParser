using BancaParser.Host.Helpers;
using DocumentFormat.OpenXml.Office2016.Excel;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Identity.Data;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using Microsoft.IdentityModel.Tokens;
using System.Collections.Generic;
using System.IdentityModel.Tokens.Jwt;
using System.Net;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace BancaParser.Host.Controllers
{
  [Route("api/[controller]")]
  [ApiController]
  public class AccountController : ControllerBase
  {
    private const string PASSWORD_HASHED = "195e6234b9b644877aacce6ff484c73c";
    private const string USERNAME_HASHED = "5d24c20cfff5958c4c68f912b79ee7bf";
    private readonly Config _configuration;

    public AccountController()
    {
      _configuration = new Config();
    }

    [HttpPost("login")]
    public IActionResult Login([FromBody] UserLogin login)
    {
      var usernameHashed = ToMd5(login.Username);
      var passwordHashed = ToMd5(login.Password);
      if (passwordHashed.ToUpper() != PASSWORD_HASHED.ToUpper() || usernameHashed.ToUpper() != USERNAME_HASHED.ToUpper())
      {
        return Unauthorized();
      }
      var secretKey = _configuration.Secret;
      var key = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(secretKey));
      var creds = new SigningCredentials(key, SecurityAlgorithms.HmacSha512);

      var claims = new[]
      {
                    new Claim(ClaimTypes.Name, login.Username),
                    new Claim(ClaimTypes.Role, "Administrator")
                };

      var token = new JwtSecurityToken(
          issuer: "banca-parser",
          audience: "banca-parser-users",
          claims: claims,
          expires: DateTime.UtcNow.AddHours(1),
          signingCredentials: creds
      );

      var jwt = new JwtSecurityTokenHandler().WriteToken(token);

      return Ok(new { token = jwt });
    }

    private static string ToMd5(string input)
    {
      using var md5 = MD5.Create();
      var inputBytes = Encoding.UTF8.GetBytes(input);
      var hashBytes = md5.ComputeHash(inputBytes);

      // Converte in stringa esadecimale
      var sb = new StringBuilder();
      foreach (var b in hashBytes)
        sb.Append(b.ToString("x2"));

      return sb.ToString();
    }
  }
  public record UserLogin(string Username, string Password);

}

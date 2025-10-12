using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Identity.Data;
using Microsoft.AspNetCore.Mvc;
using Microsoft.IdentityModel.Tokens;
using System.Collections.Generic;
using System.Net;
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

    [HttpPost("Account/Login")]
    public IActionResult Login([FromBody] LoginRequest request)
    {
      try
      {
        var usernameHashed = ToMd5(request.Email);        
        var passwordHashed = ToMd5(request.Password);
        if (passwordHashed.ToUpper() != PASSWORD_HASHED.ToUpper() || usernameHashed.ToUpper() != USERNAME_HASHED.ToUpper())
        {
          return Unauthorized("Password not valid");
        }
        
        DateTime expiresTime = DateTime.UtcNow.AddMinutes(5);
        // authentication successful so generate jwt token
        var token = JwtMiddleware.generateJwtToken(request.Email, expiresTime);
        return Ok(token);
      }
      catch (Exception ex)
      {
        return BadRequest(ex.Message);
      }
    }

    public static string ToMd5(string input)
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
}

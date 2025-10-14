using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
namespace BancaParser.Host.Controllers
{
  namespace MySecureApi.Controllers
  {
    [ApiController]
    [Route("api/[controller]")]
    public class SecureController : ControllerBase
    {
      [HttpGet("hello")]
      [Authorize]
      public IActionResult Hello()
      {
        var username = User.Identity?.Name ?? "sconosciuto";
        return Ok($"Ciao {username}, sei autenticato con successo!");
      }
    }
  }

}

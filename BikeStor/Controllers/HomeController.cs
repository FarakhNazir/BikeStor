using DocumentFormat.OpenXml.Wordprocessing;
using MailKit.Net.Smtp;
using MailKit.Security;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MimeKit;
using System.Net.Mime;

namespace BikeStor.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class HomeController : ControllerBase
    {
        [HttpPost]
        public IActionResult SendEmail( string body) 
        {
            var email = new MimeMessage();
            email.From.Add(new MailboxAddress("Farrukh","farukh.3@outlook.com"));
            email.To.Add(MailboxAddress.Parse("f2018065077@gmail.com"));
            email.Subject = "Bike store Product Sale data with total product price";
            email.Body = new TextPart(MimeKit.Text.TextFormat.Html)
            {
                Text = body
            };

            using var smtp = new SmtpClient();
            //smtp.Connect("smtp.gmail.com", 587,MailKit.Security.SecureSocketOptions.StartTls);
            smtp.Connect("smtp-mail.outlook.com", 587);
            smtp.Authenticate("farukh.3@outlook.com","bhoolgya786");
            smtp.Send(email);
            smtp.Disconnect(true);

            return Ok();
        }
        
       
    }
}

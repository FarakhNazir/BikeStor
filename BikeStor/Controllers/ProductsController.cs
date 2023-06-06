using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using BikeStor.Models;
using ClosedXML.Excel;
using MiNET.Blocks;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Reflection;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;
using MimeKit;
using Org.BouncyCastle.Utilities;
using MailKit.Net.Smtp;
using System.Net.Mail;
using System.Net.Mime;
using Microsoft.Extensions.Hosting;
using System.Net;
using LibNoise.Modifier;
using MiNET.UI;
using MailKit.Net.Smtp;
using DocumentFormat.OpenXml.InkML;
using Microsoft.AspNetCore.Http.HttpResults;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Wordprocessing;
using NuGet.Protocol.Plugins;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Options;

namespace BikeStor.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ProductsController : ControllerBase
    {
        private readonly BikeStoresContext _context;
        private object bytes;

        private IConfiguration Configuration;
        private readonly EmailSetting _emailSetting;


        public ProductsController(BikeStoresContext context, IConfiguration _configuration, IOptions<EmailSetting> mailSettingsOptions)
        {
            _context = context;
            Configuration = _configuration;
            _emailSetting = mailSettingsOptions.Value;


        }

      //  GET: api/Products
       [HttpGet("Get All Products")]
        
        public async Task <ActionResult<IEnumerable<Product >>> GetAllProducts()
        {
            if (_context.Products == null)
            {
                return NotFound();
            }

            return await _context.Products.ToListAsync();
            //SQL Query

            //SELECT p.product_id
            // ,sum(oi.quantity) AS count_of_product,
            //    p.list_price as Unit_Price,
            // sum(oi.quantity) * p.list_price as total_price
            //From production.products AS p
            //INNER JOIN sales.order_items AS oi
            //ON p.product_id = oi.product_id
            //GROUP BY p.product_id,
            // p.list_price
            //    ORDER BY p.product_id;




        }


       

       // Get Exel Sheet From Query
       [HttpGet("Get Exel List")]
        public IActionResult ExcelSheet()
        {
            var query = (from pd in _context.Products
                     join od in _context.OrderItems on pd.ProductId equals od.ProductId
                     group od by new { pd.ProductId, od.ListPrice } into results
                     orderby results.Key.ProductId
                     select new
                     {
                         ProductID = results.Key.ProductId,
                         count_of_product = results.Sum(a => a.Quantity),
                         Unit_Price = results.Key.ListPrice,
                         Sales = results.Sum(a => a.ListPrice * a.Quantity)
                     }
                   ).ToList();



            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("query");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "Product Id";
                worksheet.Cell(currentRow, 2).Value = "count_of_product";
                worksheet.Cell(currentRow, 3).Value = "Unit Price";
                worksheet.Cell(currentRow, 4).Value = "Total Sale";



                foreach (var user in query)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = user.ProductID;
                    worksheet.Cell(currentRow, 2).Value = user.count_of_product;
                    worksheet.Cell(currentRow, 3).Value = user.Unit_Price;
                    worksheet.Cell(currentRow, 4).Value = user.Sales;

                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    string filename = $"Product_{DateTime.Now.ToString("dd/mm/yyyy")}.xlsx";
                    return File(
                              content,
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                  "Product_Sale_Detail.xlsx");


                }
              
            }
        }

        [HttpGet("Send Exel sheet Send via email Automaticaly")]
        public IActionResult SendEmailWithQuryGenratedExelFile(string sendto)
        {
            var query = (from pd in _context.Products
                     join od in _context.OrderItems on pd.ProductId equals od.ProductId
                     group od by new { pd.ProductId,od.ListPrice } into results
                     orderby results.Key.ProductId
                     select new
                     {
                         ProductID = results.Key.ProductId,
                         count_of_product = results.Sum(a => a.Quantity),
                         Unit_Price = results.Key.ListPrice,
                         Sales = results.Sum(a => a.ListPrice * a.Quantity)
                     }
                   ).ToList();



            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("query");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "Product Id";
                worksheet.Cell(currentRow, 2).Value = "count_of_product";
                worksheet.Cell(currentRow, 3).Value = "Unit Price";
                worksheet.Cell(currentRow, 4).Value = "Total Sale";



                foreach (var user in query)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = user.ProductID;
                    worksheet.Cell(currentRow, 2).Value = user.count_of_product;
                    worksheet.Cell(currentRow, 3).Value = user.Unit_Price;
                    worksheet.Cell(currentRow, 4).Value = user.Sales;

                }

              

                    System.IO.MemoryStream theStream = new System.IO.MemoryStream();
                    workbook.SaveAs(theStream);
                   
                    byte[] byteArr = theStream.ToArray();
                    System.IO.MemoryStream stream1 = new System.IO.MemoryStream(byteArr, true);
                    stream1.Write(byteArr, 0, byteArr.Length);
                    stream1.Position = 0;

                //Read SMTP settings from AppSettings.json.
               
                /*string host = _emailSetting.PrimaryDomain;
                int port = _emailSetting.PrimaryPort;
                string fromAddress = _emailSetting.FromEmail;
                string userName = _emailSetting.UsernameEmail;
                string password = _emailSetting.UsernamePassword;

                MailMessage message = new MailMessage(
                           fromAddress,
                            sendto,
                             "Product Sale Detail with Price List.",
                             " See the attached spreadsheet.");


                message.Attachments.Add(new Attachment(stream1, "Produt_Total_Sale_Detail.xlsx"));
                System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient(host, 587);
                smtp.UseDefaultCredentials = false;
                NetworkCredential nc = new NetworkCredential(userName, password);
                smtp.Credentials = nc;
                smtp.EnableSsl = true;
                smtp.Send(message);*/

                MailMessage message = new MailMessage(
                            "farukh.3@outlook.com",
                            sendto,
                             "Product Sale Detail with Price List.",
                             " See the attached spreadsheet.");


                message.Attachments.Add(new Attachment(stream1, "Produt_Total_Sale_Detail.xlsx"));
                System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient("smtp-mail.outlook.com", 587);
                smtp.UseDefaultCredentials = false;
                NetworkCredential nc = new NetworkCredential("farukh.3@outlook.com", "bhoolgya786");
                smtp.Credentials = nc;
                smtp.EnableSsl = true;
                smtp.Send(message);

                return Ok();


               
            }


        }
                        

                       
                     
[HttpGet("Send Email With Exel sheet")]
public IActionResult SendEmail(string sendto)
{
    string exelfile = "Product_Sale_Detail_2.xlsx";
    Attachment data = new Attachment(exelfile, MediaTypeNames.Application.Octet);
    System.Net.Mime.ContentDisposition disposition = data.ContentDisposition;
    disposition.CreationDate = System.IO.File.GetCreationTime(exelfile);
    disposition.ModificationDate = System.IO.File.GetLastWriteTime(exelfile);
    disposition.ReadDate = System.IO.File.GetLastAccessTime(exelfile);
    MailMessage message = new MailMessage(
            "farukh.3@outlook.com",
            sendto,
             "Product Sale Detail with Price List.",
             " See the attached spreadsheet.");

    message.Attachments.Add(data);
    System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient("smtp-mail.outlook.com", 587);
    smtp.UseDefaultCredentials = false;
    NetworkCredential nc = new NetworkCredential("farukh.3@outlook.com", "bhoolgya786");
    smtp.Credentials = nc;
    smtp.EnableSsl = true;
    smtp.Send(message);

    return Ok();

}
//GET: api / Products / 5
[HttpGet("{id}")]
        public async Task<ActionResult<Product>> GetProduct(int id)
{
    if (_context.Products == null)
    {
        return NotFound();
    }
    var product = await _context.Products.FindAsync(id);

    if (product == null)
    {
        return NotFound();
    }

    return product;
}

// PUT: api/Products/5
// To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
[HttpPut("{id}")]
public async Task<IActionResult> PutProduct(int id, Product product)
{
    if (id != product.ProductId)
    {
        return BadRequest();
    }

    _context.Entry(product).State = EntityState.Modified;

    try
    {
        await _context.SaveChangesAsync();
    }
    catch (DbUpdateConcurrencyException)
    {
        if (!ProductExists(id))
        {
            return NotFound();
        }
        else
        {
            throw;
        }
    }

    return NoContent();
}

// POST: api/Products
// To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
[HttpPost]
public async Task<ActionResult<Product>> PostProduct(Product product)
{
    if (_context.Products == null)
    {
        return Problem("Entity set 'BikeStoresContext.Products'  is null.");
    }
    _context.Products.Add(product);
    await _context.SaveChangesAsync();

    return CreatedAtAction("GetProduct", new { id = product.ProductId }, product);
}

// DELETE: api/Products/5
[HttpDelete("{id}")]
public async Task<IActionResult> DeleteProduct(int id)
{
    if (_context.Products == null)
    {
        return NotFound();
    }
    var product = await _context.Products.FindAsync(id);
    if (product == null)
    {
        return NotFound();
    }

    _context.Products.Remove(product);
    await _context.SaveChangesAsync();

    return NoContent();
}

private bool ProductExists(int id)
{
    return (_context.Products?.Any(e => e.ProductId == id)).GetValueOrDefault();
}


    }
}


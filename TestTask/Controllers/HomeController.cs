using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using TestTask.Models;
using System.Net.Mail;
using System.IO;
using System.Data;
using OfficeOpenXml;

namespace TestTask.Controllers
{
    public class HomeController : Controller
    {
        string filename = "SalesReport.xlsx";

        UserContext db = new UserContext();

        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public ActionResult Form()
        {
            bool reportGenerated = false;
            ViewBag.isGen = reportGenerated;
           
            return View();

        }

        [HttpPost]
        public ViewResult Form(DateTime Date1, DateTime Date2)
        {
            bool reportGenerated = true;
            string filepath = Server.MapPath($"~/Reports/{filename}");
            var fileInfo = new FileInfo(filepath);

            IEnumerable<OrderDetail> order_details = db.OrderDetails
                .Where(p => (p.Order.OrderDate > Date1 && p.Order.OrderDate <= Date2));
            if (fileInfo.Exists)
                fileInfo.Delete();

            var xls = new ExcelPackage(fileInfo);
            
            var sheet = xls.Workbook.Worksheets.Add("Sales report");

                sheet.Cells[1, 1].Value = "Код заказа";
                sheet.Cells[1, 2].Value = "Дата заказа";
                sheet.Cells[1, 3].Value = "Код продукта";
                sheet.Cells[1, 4].Value = "Название продукта";
                sheet.Cells[1, 5].Value = "Продано";
                sheet.Cells[1, 6].Value = "Цена за штуку";
                sheet.Cells[1, 7].Value = "Общая стоимость";

            var i = 1;
                foreach (var att in order_details)
                {
                    i++;

                    sheet.Cells[i, 1].Value = att.OrderID;
                    sheet.Cells[i, 2].Value = att.Order.OrderDate;
                    sheet.Cells[i, 3].Value = att.ProductID;
                    sheet.Cells[i, 4].Value = att.Product.Name;
                    sheet.Cells[i, 5].Value = att.Quantity;
                    sheet.Cells[i, 6].Value = att.UnitPrice;
                    sheet.Cells[i, 7].Formula = $"=E{i}*(F{i}";
            }
            sheet.Cells.AutoFitColumns();
            xls.Save();
            

            

            ViewBag.isGen = reportGenerated;
            ViewBag.OrderDetail = order_details;
            ViewBag.firstDate = Date1.ToString("dd-MM-yyyy");
            ViewBag.secondDate = Date2.ToString("dd-MM-yyyy");
            return View();
        }

        
        [HttpGet]
        public ActionResult Report( )
        {
            
            return View();
        }

        [HttpPost]
        public ViewResult Report(string Email)
        {
            string filepath = Server.MapPath($"~/Reports/{filename}");
            var fileInfo = new FileInfo(filepath);

            try
            {
                if (ModelState.IsValid)
                {
                    var senderEmail = new MailAddress("Kup9Python@gmail.com", "Sender");
                    var recieverEmail = new MailAddress(Email, "Reciever");
                    var password = "djangodev";
                    var subject = "Your sales report";
                    var body = "You clamed sales report from Northwind";
                    var smtp = new SmtpClient
                        {
                            Host = "smtp.gmail.com",
                            Port = 587,
                            EnableSsl = true,
                            DeliveryMethod = SmtpDeliveryMethod.Network,
                            UseDefaultCredentials = false,
                            Credentials = new System.Net.NetworkCredential(senderEmail.Address, password),

                        };
                    

                    using (var message = new MailMessage(senderEmail, recieverEmail)
                    {
                        Subject = subject,
                        Body = body
                    })
                    {
                        if (fileInfo.Exists)
                        {
                            message.Attachments.Add(new Attachment(filepath));
                            smtp.Send(message); }
                        else
                        {
                            ViewBag.Error = "Какие то проблемы с созданием отчета. Попробоуйте еще раз или обратитесь к модератору за помощью";
                            return View();
                        }
                    }
                }
            }
            catch (Exception)
            {
                ViewBag.Error = "Какие то проблемы с отправкой письма. Попробоуйте еще раз или обратитесь к модератору за помощью";
                return View();
            }

            return View("Success");
        }
    }
}
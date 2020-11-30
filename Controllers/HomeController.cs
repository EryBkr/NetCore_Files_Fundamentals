using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using FileExample.Models;
using OfficeOpenXml;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Data;
using FastMember;

namespace FileExample.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        public FileContentResult GetExcelFile()
        {
            ExcelPackage package = new ExcelPackage();
            var blank=package.Workbook.Worksheets.Add("Sayfa1");//Excel sayfasını belirledik

            blank.Cells[1, 1].Value = "Name"; //Excel hücrelerine değerleri verdik
            blank.Cells[1, 2].Value = "Soyad";
            blank.Cells[2, 1].Value = "Eray";
            blank.Cells[2, 2].Value = "Bakır";

           var byteExcel= package.GetAsByteArray(); //Dosyamızı byte dizisine çevirdik
            
            return File(byteExcel, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", Guid.NewGuid()+ "" + ".xlsx"); //Byte dizisi Content Type ve Dosya ismini verdik
        }

        public FileContentResult GetExcelFileWithObject()
        {
            ExcelPackage package = new ExcelPackage();
            var blank = package.Workbook.Worksheets.Add("Musteri");//Excel sayfasını belirledik
            blank.Cells["A1"].LoadFromCollection(new List<Person> { new Person { Id = 1, Name = "Eray" }, new Person { Id = 2, Name = "Berkay" } },true,OfficeOpenXml.Table.TableStyles.Dark1); //Sütuna yazılacak listeyi ,Header bilgisini istiyorsak header bilgisini ve Theme bilgisini veriyoruz
           

            var byteExcel = package.GetAsByteArray(); //Dosyamızı byte dizisine çevirdik

            return File(byteExcel, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", Guid.NewGuid() + "" + ".xlsx"); //Byte dizisi Content Type ve Dosya ismini verdik
        }

        public IActionResult GetPDF()
        {
            Document document = new Document(PageSize.A4,25f,25f,25f,25f); //Sayfa boyutunu ve boşlukları belirledik

            string fileName = Guid.NewGuid() + ".pdf";

            string path = Path.Combine(Directory.GetCurrentDirectory(),"wwwroot/documents/"+fileName);

            var stream = new FileStream(path, FileMode.Create);
            PdfWriter.GetInstance(document, stream);


            document.Open();
            Paragraph paragraph = new Paragraph(new Phrase("Eray Bakır"));
            document.Add(paragraph);


            document.Close();

            return File("/documents/" + fileName, "application/pdf", fileName);
        }

        public IActionResult GetPDFWithTable()
        {
            Document document = new Document(PageSize.A4, 25f, 25f, 25f, 25f); //Sayfa boyutunu ve boşlukları belirledik

            string fileName = Guid.NewGuid() + ".pdf";

            string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/documents/" + fileName);

            var stream = new FileStream(path, FileMode.Create);
            PdfWriter.GetInstance(document, stream);


            document.Open();

            PdfPTable pdfPTable = new PdfPTable(2); //Kolon sayısı
            pdfPTable.AddCell("Ad");
            pdfPTable.AddCell("Soyad");

            pdfPTable.AddCell("Eray");
            pdfPTable.AddCell("Bakır");

            document.Add(pdfPTable);


            document.Close();

            return File("/documents/" + fileName, "application/pdf", fileName);
        }

        public IActionResult GetPDFWithObject()
        {
            Document document = new Document(PageSize.A4, 25f, 25f, 25f, 25f); //Sayfa boyutunu ve boşlukları belirledik

            DataTable dataTable = new DataTable();
            dataTable.Load(ObjectReader.Create(new List<Person> { new Person { Id = 1, Name = "Eray" }, new Person { Id = 2, Name = "Berkay" } }));

            string fileName = Guid.NewGuid() + ".pdf";

            string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/documents/" + fileName);

            var stream = new FileStream(path, FileMode.Create);
            PdfWriter.GetInstance(document, stream);


            document.Open();

            PdfPTable pdfPTable = new PdfPTable(dataTable.Columns.Count); //Kolon sayısı
            
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                pdfPTable.AddCell(dataTable.Columns[i].ColumnName);
            }

            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    pdfPTable.AddCell(dataTable.Rows[i][j].ToString());
                }
            }

            document.Add(pdfPTable);


            document.Close();

            return File("/documents/" + fileName, "application/pdf", fileName);
        }

    }

    class Person
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }
}

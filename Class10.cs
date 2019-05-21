using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.IO;

namespace LSDemo
{
    class Class10
    {
        public static void CreateNewPdf()
        {
            try
            {
                PdfReader reader = new PdfReader(@"d:\P19D41856A-new-drawings-20190513-wzi.pdf");
                int pageNums = reader.NumberOfPages;
                Document doc = new Document(PageSize.A4, 0, 0, 0, 0);
                PdfWriter writer = PdfWriter.GetInstance(doc,
                new FileStream(@"D:\new.pdf",
                FileMode.Create));
                doc.Open();
                PdfContentByte cb = writer.DirectContent;
                for (int i = 1; i <= pageNums; i++)
                {
                    PdfImportedPage page = writer.GetImportedPage(reader, i); //page #1
                    float Scale = 1.0f;
                    cb.AddTemplate(page, Scale, 0, 0, Scale, 0, 0);
                    doc.NewPage();
                }
                doc.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}

using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO.Enumeration;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;

namespace common{
    public class GuidePdf{
        public Boolean createPDFReport(String Guide, ref StringBuilder attachments) {
            iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance("https://pidoboxmiami.com/wp-content/uploads/2019/03/Logo-Pidobox-LLC.png");
            img.ScaleToFit(200, 100);
            img.SetAbsolutePosition(10, 760);

            String body = Properties.Resources.TemplateFileMail.ToString();
            Console.WriteLine(body);
            using (FileStream stream = new FileStream("C:\\Users\\yuno\\Desktop\\angular.pdf", FileMode.Create)) {
                iTextSharp.text.Document Pdf = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 50, 50, 50, 50);

                PdfWriter writer = PdfWriter.GetInstance(Pdf, stream);

                Pdf.Open();
                Pdf.Add(new Phrase(""));
                Pdf.Add(img);

                using (StringReader sr = new StringReader(body)) {
                    XMLWorkerHelper.GetInstance().ParseXHtml(writer, Pdf, sr);
                }

                Pdf.Close();
                stream.Close();
            }

            return true;
        }
    }
}

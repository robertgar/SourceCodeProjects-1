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
            String body = "";
            Console.WriteLine(body);
            using (FileStream stream = new FileStream("C:\\Users\\yuno\\Desktop\\angular.pdf", FileMode.Create)) {
                iTextSharp.text.Document Pdf = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 50, 50, 50, 50);

                PdfWriter writer = PdfWriter.GetInstance(Pdf, stream);

                Pdf.Open();
                Pdf.Add(new Phrase(""));

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

using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using System.Data;
using System.Text;

namespace common {
    public class GuidePdf {
        public Boolean createPDFReport(DataRow Row, ref StringBuilder attachments, ref String Error) {
            try {
                createPDF(ref Row, ref attachments);
                return true;
            } catch (Exception e) {
                Error = e.ToString().Substring(0, 50);
                return false;
            }
        }

        private void createPDF(ref DataRow row, ref StringBuilder attachments) {
            iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance("http://yocargo.com/wp-content/uploads/2019/12/LogoYoCargo.jpg");
            StringBuilder body = new StringBuilder();
            body.Append(Properties.Resources.TemplateFileMail.ToString());

            switch (row["CodigoImportador"].ToString()) {
                case "6":
                    img = iTextSharp.text.Image.GetInstance("https://pidoboxmiami.com/wp-content/uploads/2019/03/Logo-Pidobox-LLC.png");
                    body.Replace("@GuideCode", "Pobox guide<br />" + row["GuiaAerea"].ToString());
                    break;
                case "7":
                    img = iTextSharp.text.Image.GetInstance("https://static.wixstatic.com/media/fa8e50_cb857d1789fd41c1bcef2222b8500c7b~mv2.png/v1/fill/w_462,h_98,al_c,usm_0.66_1.00_0.01/PRESSEX%20COLOR.png");
                    break;
                case "8":
                    img = iTextSharp.text.Image.GetInstance("https://quickshipping.com/_assets/imgs/logos/logo_quick_shipping_logistics.png");
                    body.Replace("@GuideCode", "QS guide<br />" + row["GuiaAerea"].ToString());
                    break;
                case "9":
                    img = iTextSharp.text.Image.GetInstance("https://icclogistic.com/wp-content/uploads/2022/07/ICC-LOGISTIC-ForWhiteBack-1-1024x311.png");
                    body.Replace("@GuideCode", "ICC guide<br />" + row["GuiaAerea"].ToString());
                    break;
                default:
                    body.Replace("@GuideCode", "Guide<br />" + row["GuiaAerea"].ToString());
                    break;
            }

            img.ScaleToFit(200, 100);
            img.SetAbsolutePosition(10, 760);

            body.Replace("@Dress", row["Awb"].ToString());
            body.Replace("@CompanyName", row["Nombre"].ToString());
            body.Replace("@PhoneCompany", "");
            body.Replace("@Date", row["FechaEnvio"].ToString());
            body.Replace("@Origin", "");
            body.Replace("@AirNumber", "");
            body.Replace("@Destination", "Guatemala");
            body.Replace("@Description", row["Descripcion"].ToString());
            body.Replace("@Parts", row["Parts"].ToString());
            body.Replace("@Weigth", row["Peso"].ToString());
            body.Replace("@Tracking", row["CodigoDeRastreo"].ToString());
            body.Replace("@Charge", row["Charge"].ToString());

            using (FileStream stream = new FileStream("C:\\inetpub\\wwwroot\\Sistema\\CARGA\\" + row["CodigoDeRastreo"].ToString() + ".pdf", FileMode.Create)) {
                iTextSharp.text.Document Pdf = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 50, 50, 50, 50);

                PdfWriter writer = PdfWriter.GetInstance(Pdf, stream);

                Pdf.Open();
                Pdf.Add(new Phrase(""));
                Pdf.Add(img);

                using (StringReader sr = new StringReader(body.ToString())) {
                    XMLWorkerHelper.GetInstance().ParseXHtml(writer, Pdf, sr);
                }

                Pdf.Close();
                stream.Close();
            }

            if (attachments.ToString().Contains("C:\\inetpub\\wwwroot\\Sistema\\CARGA\\" + row["CodigoDeRastreo"].ToString() + ".pdf")) { return; }
            if (attachments.Length > 0) { attachments.Append(","); }

            attachments.Append("C:\\inetpub\\wwwroot\\Sistema\\CARGA\\" + row["CodigoDeRastreo"].ToString() + ".pdf");
        }
    }
}

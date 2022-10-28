using System.Text;
using System.Net.Mail;
using System.Net;

namespace connection{
    public class Data {
        public Boolean isSimulation;
        public StringBuilder origin = new StringBuilder();
        public StringBuilder destination = new StringBuilder();
        public StringBuilder attachments = new StringBuilder();
        public StringBuilder subject = new StringBuilder();
        public StringBuilder body = new StringBuilder();
        public StringBuilder cc = new StringBuilder();
        public void clearAll(){
            origin.Clear();
            destination.Clear();
            attachments.Clear();
            subject.Clear();
            body.Clear();
            cc.Clear();
        }
    }
    public class SendEmails {
        public Data data = new Data();
        private common.UseCommon? use;
        private Execute? execute;
        public void setSimulation(ref common.UseCommon use, ref Execute execute) {
            this.use = use;
            this.execute = execute;
        }

        public Boolean sendEmail(ref String Error) {
            try {
                trySendEmail();
                return true;
            } catch (Exception e) {
                Error = e.ToString();
                return false;
            }
        }
        
        public void trySendEmail() {
            if (data.origin.ToString().Trim().Equals("")) {
                Console.WriteLine("Cannon send mail because Origin is empty.");
                return;
            }
            if (data.destination.ToString().Trim().Equals("")) {
                Console.WriteLine("Cannot send mail because Destination is empty");
                return;
            }
            if (data.subject.ToString().Trim().Equals("")) {
                Console.WriteLine("Cannot send mail because Title is empty.");
                return;
            }
            if (data.body.ToString().Trim().Equals("")) {
                Console.WriteLine("Cannot send mail because Body is empty.");
                return;
            }

            data.body.Replace(data.body.ToString(), connection.Properties.Resources.TemplatePrincipalMail.ToString().Replace("@body", data.body.ToString()));

            using (MailMessage mail = new MailMessage(new MailAddress(data.origin.ToString(), "GuatemalaDigital.com"), new MailAddress(data.destination.ToString()))) {
                foreach (String cc in data.cc.ToString().Split(",")) {
                    if (cc.Trim().Equals("")) { continue; }
                    mail.CC.Add(cc);
                }
                foreach (String file in data.attachments.ToString().Split(",")) {
                    if (file.Trim().Equals("")) { continue; }
                    mail.Attachments.Add(new Attachment(file));
                }

                mail.Subject = data.subject.ToString();
                mail.Body = data.body.ToString();
                mail.IsBodyHtml = true;
                mail.Priority = MailPriority.Normal;
                using (SmtpClient sender = new SmtpClient()) {
                    sender.UseDefaultCredentials = false;
                    sender.Host = "email-smtp.us-east-1.amazonaws.com";
                    sender.Port = 587;
                    sender.EnableSsl = true;
                    sender.Credentials = new NetworkCredential("AKIAZJ6ZLR4SSIO3C2X6", "BCuKYZn5qkcXIFT5YNSxgVu0uQejzt+w4q8+F6hUIROW");
                    sender.Send(mail);
                }
            }
        }
    }
}

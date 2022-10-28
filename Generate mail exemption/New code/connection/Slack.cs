using System.Net;
using System.Text;

namespace connection{
    public class Slack {
        public Data data = new Data();
        public class Data {
            public StringBuilder subject = new StringBuilder();
            public StringBuilder chanel = new StringBuilder();
            public StringBuilder alertTitle = new StringBuilder();
            public StringBuilder message = new StringBuilder();
            public StringBuilder procedure = new StringBuilder();

            public void Clear() {
                subject.Clear();
                chanel.Clear();
                alertTitle.Clear();
                message.Clear();
                procedure.Clear();
            }
        }
        
        public void send() {
            if (data.subject.ToString().Trim().Equals("")) { return; }
            if (data.chanel.ToString().Trim().Equals("")) { return; }
            if (data.alertTitle.ToString().Trim().Equals("")) { return; }
            if (data.message.ToString().Trim().Equals("")) { return; }
            if (data.procedure.ToString().Trim().Equals("")) { return; }

            data.subject.Append("\nServer: ").Append(Environment.MachineName);

            data.message.Replace(data.message.ToString(), Properties.Resources.structureSlack.ToString().Replace("@Message", data.message.ToString()));
            data.message.Replace("@Subject", data.subject.ToString());
            data.message.Replace("@AlertTitle", data.alertTitle.ToString());
            data.message.Replace("@Procedure", data.procedure.ToString());
            
            try{
                trySend();
            }catch(Exception e) {
                Console.WriteLine(e);
            }
        }
        private void trySend() {
            WebClient web = new WebClient();
            web.Headers[HttpRequestHeader.ContentType] = "application/json; chartset=utf-8";
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            web.UploadString(data.chanel.ToString(), data.message.ToString());
        }
    }
}
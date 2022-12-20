using System.Net;
using System.Text;

namespace connection {
    public class Slack {
        public Data data = new Data();
        public class Data {
            public StringBuilder subject = new StringBuilder();
            public StringBuilder alertTitle = new StringBuilder();
            public StringBuilder message = new StringBuilder();
            public string warningColour = "";
            public typeAlertColor selectColour = new typeAlertColor();
            
            public void Clear() {
                subject.Clear();
                alertTitle.Clear();
                message.Clear();
                warningColour = "";
            }

            public struct typeAlertColor {
                public readonly string Red = "#e60000";
                public readonly string Green = "#00b300";
                public readonly string Blue = "#0073e6";

                public typeAlertColor() {}
            }
        }

        public void sendError(string Error) {
            data.Clear();
            data.alertTitle.Append("Error alert");
            data.subject.Append("Procedure alert inf: Ops! Something has gone wrong. :(");
            data.warningColour = data.selectColour.Red;
            data.message.Append(Error.Replace(((char)34).ToString(), "'"));
            send();
        }
        
        public void send() {
            if (data.subject.ToString().Trim().Equals("")) { return; }
            if (data.alertTitle.ToString().Trim().Equals("")) { return; }
            if (data.message.ToString().Trim().Equals("")) { return; }
            if (data.warningColour.Equals("")) { data.warningColour = data.selectColour.Blue; }

            try {
                wsGD.ServiceSoapClient _slack = new wsGD.ServiceSoapClient();
                _slack.EnviarAlertaSlack(
                    data.message.ToString(), 
                    data.subject.ToString(), 
                    "Generate mail exemtion",
                    11,
                    data.alertTitle.ToString(),
                    data.warningColour
                    );
            } catch (Exception e) {
                Console.WriteLine(e);
            }
        }
    }
}
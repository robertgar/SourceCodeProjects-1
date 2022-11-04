namespace Principal {
    internal class Maker {
        internal class Vars {
            public connection.Slack slack = new connection.Slack();
            public connection.SQL execute = new connection.SQL();
            public common.UseCommon use = new common.UseCommon();
        }
        public void MakeAll(ref Boolean isSimulation) {
            Vars js = new Vars();
            DateTime _now = DateTime.Now;
            
            js.slack.data.Clear();
            js.slack.data.chanel.Append(js.execute.getChanel(11));
            js.slack.data.alertTitle.Append("Alert");
            js.slack.data.subject.Append("Procedure alert info: the process has begun...");
            js.slack.data.procedure.Append("Procedure: charge ads");
            js.slack.data.message.Append("Start time: ").Append(_now);
            js.slack.send();

            try {
                TryMake(ref js);
            } catch(Exception e) {
                js.slack.sendError(e.ToString(), js.execute.getChanel(11));
            }

            js.slack.data.Clear();
            js.slack.data.chanel.Append(js.execute.getChanel(11));
            js.slack.data.alertTitle.Append("Alert");
            js.slack.data.subject.Append("Procedure alert info: the process has been completed!");
            js.slack.data.procedure.Append("Procedure: charge ads");
            js.slack.data.message.Append("Total execution time: ").Append(DateTime.Now - _now);
            js.slack.send();
        }
        
        private void TryMake(ref Vars js) {
            
        }
    }
}
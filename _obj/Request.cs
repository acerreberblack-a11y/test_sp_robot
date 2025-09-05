using System;
using System.Collections.Generic;

namespace SpravkoBot_AsSapfir
{
    internal class Request
    {
        public string RequestType { get; set; }
        public string RequestBE { get; set; }
        public string RequestUUID { get; set; }
        public DateTime DateStart { get; set; }
        public DateTime DateEnd { get; set; }
        public string INN { get; set; }
        public string KPP { get; set; }
        public string Service { get; set; }
        public string Organization { get; set; }
        public List<string> AgreementNumbers { get; set; }
        public List<string> CounterpartyNumbers { get; set; }
        public string message { get; set; }
        public string status { get; set; }
        public Queue<SapTask> TaskQueue { get; } = new Queue<SapTask>();
    }
}

using System;
using System.Collections.Generic;

namespace SpravkoBot_AsSapfir
{
    internal class SapTask
    {
        public string Organization { get; set; }
        public string BeCode { get; set; }
        public string INN { get; set; }
        public string KPP { get; set; }
        public List<string> AgreementNumbers { get; set; }
        public List<string> CounterpartyNumbers { get; set; }
        public DateTime DateStart { get; set; }
        public DateTime DateEnd { get; set; }
    }
}

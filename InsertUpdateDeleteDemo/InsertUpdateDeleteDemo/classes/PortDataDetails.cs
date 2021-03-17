using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsertUpdateDeleteDemo.classes
{
    class PortDataDetails
    {
        public string partNumber { get; set; }
        public string revisionNumber { get; set; }
        public string dateTime { get; set; }
        public string testResult { get; set; }
        public string testingMode { get; set; }
        public string passCount { get; set; }
        public string failCount { get; set; }
        public string failTypePoints { get; set; }
        public string cutterCount { get; set; }
        public string labelSerialNumber { get; set; }
        public string printedBarCode { get; set; }
        public string operatorCode { get; set; }
        public string shift { get; set; }
        public string lotCount { get; set; }
        public string lotQuantity { get; set; }
        public string customField { get; set; }
    }
}

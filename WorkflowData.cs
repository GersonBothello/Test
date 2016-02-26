using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication1
{
    public class WorkflowData
    {
        public string name { get; set; }
        public string active { get; set; }
        public string owner { get; set; }
        public string creator { get; set; }
        public DateTime activationtime { get; set; }
        public string description { get; set; }
        public string JobId { get; set; }
        public string lasttask { get; set; }
        public string TaskDescription { get; set; }
        public DateTime LastModified { get; set; }
        public int DaysLastModified { get; set; }
    }
}

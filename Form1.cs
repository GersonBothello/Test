using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

New branch 3 stahs testrttt
        const string folderPath = @"C:\_Projects\WindowsFormsApplication1\Active_workflows\Active_workflows\";

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string folderPath = txtFolderPath.Text;

                if (!string.IsNullOrEmpty(folderPath))
                {
                    GenerateExcelFile createFiles = new GenerateExcelFile();
                   var errorFiles = createFiles.CreateFiles(folderPath);
                   txtDestination.Text = folderPath + "\\" + "workflow.csv";


                   string errorFile = string.Empty;
                   foreach (string item in errorFiles)
                   {
                       errorFile = item + System.Environment.NewLine;
                   }

                   txtError.Text = errorFile;
                }

                MessageBox.Show("Complete");
            
            }
            catch (Exception ex)
            {
                txtError.Text = ex.Message;
            }
        }

        private void GetData()
        {

            var xmlFiles = System.IO.Directory.GetFiles(folderPath + "\\fixFiles\\");
            List<WorkflowData> workflowList = new List<WorkflowData>();

            foreach (string file in xmlFiles)
            {
                GeFileData(file, workflowList);
            }

            WriteWorkFlowFile(workflowList);

        }

        private void FixFile()
        {

            var xmlFiles = System.IO.Directory.GetFiles(folderPath);

            foreach (string file in xmlFiles)
            {
                fixXml(file);
            }

        }


        private void fixXml(string filePath)
        {
            List<string> writeLines = new List<string>();

            var fileName = System.IO.Path.GetFileName(filePath);

            using (System.IO.StreamReader reader = new System.IO.StreamReader(filePath))
            {
                while (!reader.EndOfStream)
                {
                    string line = reader.ReadLine();

                    if (line != "</workflow>")
                    {
                        writeLines.Add(line);
                    }
                }

                writeLines.Add("</workflow>");
            }

            if (!System.IO.Directory.Exists(folderPath + "\\fixFiles\\"))
            {
                System.IO.Directory.CreateDirectory(folderPath + "\\fixFiles\\");
            }

            using (System.IO.StreamWriter writer = new System.IO.StreamWriter(folderPath + "\\fixFiles\\" + fileName))
            {
                foreach (string item in writeLines)
                {
                    writer.WriteLine(item);
                }
            }
        }




        public string GeFileData(string filePath, List<WorkflowData> workflowList)
        {

            string returnValue = string.Empty;

            List<WorkflowEventData> eventList = new List<WorkflowEventData>();

            System.Xml.Linq.XDocument xDoc = System.Xml.Linq.XDocument.Load(filePath);

            var result = xDoc.Descendants("workflow");


            foreach (var item in result)
            {
                var name = item.Attribute("name").Value;
                var active = item.Attribute("active").Value;
                var owner = item.Attribute("owner").Value;
                var creator = item.Attribute("creator").Value;
                var activationtime = FromUnixTime(item.Attribute("activationtime").Value);
                var description = item.Descendants("description").First().Value;

                var jobId = item.Descendants("jobid").First().Value;
                var row = new WorkflowData();

                var lastTask = item.Descendants("lasttask").First().Value;

                string[] splitLastTask = lastTask.Split('\n');

                foreach (string taskItem in splitLastTask)
                {
                    string[] colonSplit = taskItem.Split(':');

                    if (colonSplit[0].Equals("Taskname", StringComparison.CurrentCultureIgnoreCase))
                    {
                        row.lasttask = colonSplit[1];
                    }

                    if (colonSplit[0].Equals("Description", StringComparison.CurrentCultureIgnoreCase))
                    {
                        row.TaskDescription = colonSplit[1];
                    }

                }

                if (creator.Equals("yc19d1s", StringComparison.CurrentCultureIgnoreCase) || creator.Equals("w003spj", StringComparison.CurrentCultureIgnoreCase))
                {
                    continue;
                }

                row.name = name;
                row.active = active;
                row.owner = owner;
                row.creator = creator;
                row.activationtime = activationtime;
                row.description = description;
                row.JobId = jobId;
                //row.lasttask = lastTask;

                var workflowEvent = item.Descendants("event");

                foreach (var eventItem in workflowEvent)
                {
                    WorkflowEventData eventrow = new WorkflowEventData();

                    eventrow.date = FromUnixTime(eventItem.Attribute("date").Value);
                    eventrow.name = eventItem.Attribute("name").Value;
                    eventrow.user = eventItem.Attribute("user").Value;
                    eventrow.rolename = eventItem.Attribute("rolename").Value;
                    eventrow.task = eventItem.Attribute("task").Value;
                    eventrow.taskname = eventItem.Attribute("taskname").Value;
                    eventrow.other = eventItem.Attribute("other").Value;
                    eventrow.othername = eventItem.Attribute("othername").Value;

                    eventList.Add(eventrow);

                }

                var modifiedDate = (from d in eventList select d.date).Max();

                var lastEventRow = eventList.LastOrDefault();

                row.LastModified = modifiedDate;

                DateTime startTime = Convert.ToDateTime("01/06/2016");

                row.DaysLastModified = (startTime - modifiedDate.Date).Days;

                workflowList.Add(row);
            }




            return returnValue;
        }


        private void WriteWorkFlowFile(List<WorkflowData> listData)
        {
            System.IO.StreamWriter writer = new System.IO.StreamWriter(folderPath + "workflow.csv");
            string comma = ",";

            string header = "name" + comma + "active" + comma + "owner" + comma + "creator" + comma + "activationtime" + comma + "description" + comma + "Job Id" + comma + "Last Modified" + comma + "Last Task" + comma + "Task Description" + comma + "Days Since Last Modified " + DateTime.Today.ToString();
            writer.WriteLine(header);

            foreach (WorkflowData item in listData)
            {
                string row = item.name + comma + item.active + comma + item.owner + comma + item.creator + comma + item.activationtime.ToString() + comma + item.description + comma + item.JobId + comma + item.LastModified + comma + item.lasttask + comma + item.TaskDescription + comma + item.DaysLastModified;
                writer.WriteLine(row);
            }

            writer.Close();
            writer.Dispose();
        }


        private void WriteEventFile(List<WorkflowEventData> listData)
        {
            System.IO.StreamWriter writer = new System.IO.StreamWriter(@"C:\Personal\Macys\Active Jobs\Active Jobs\workflowEvent.csv");
            string comma = ",";

            string header = "date" + comma + "name" + comma + "user" + comma + "rolename" + comma + "task" + comma + "taskname" + comma + "other" + comma + "othername";
            writer.WriteLine(header);

            foreach (WorkflowEventData item in listData)
            {
                string row = item.date.ToString() + comma + item.name + comma + item.user + comma + item.rolename + comma + item.task + comma + item.taskname + comma + item.other + comma + item.othername;
                writer.WriteLine(row);
            }

            writer.Close();
            writer.Dispose();
        }


        public DateTime FromUnixTime(string unixTimeString)
        {

            long unixTime = long.Parse(unixTimeString);

            var epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            return epoch.AddSeconds(unixTime);
        }

        public static string GetRootImgAlt(System.Xml.Linq.XDocument xDoc, string nodeName)
        {
            string returnValue = string.Empty;
            var result = xDoc.Descendants(nodeName);

            if (result != null)
            {
                if (result.Count() > 0)
                {
                    XElement element = result.First();
                    if (element != null)
                    {
                        if (element.HasElements)
                        {
                            //returnValue = element.Descendants(Img).Attributes(Alt).First().Value;
                        }
                    }
                }
            }
            return returnValue;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowser = new FolderBrowserDialog();

            folderBrowser.ShowDialog();

            string folderPath = folderBrowser.SelectedPath;

            txtFolderPath.Text = folderPath;
            
            //if (!string.IsNullOrEmpty(folderPath))
            //{
            //    GenerateExcelFile createFiles = new GenerateExcelFile();
            //    createFiles.CreateFiles(folderPath);
            //}
        }

    }
}

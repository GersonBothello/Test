using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication1
{
    class GenerateExcelFile
    {

        List<string> ErrorFileList = new List<string>();

        public List<string> CreateFiles(string folderPath)
        {

            string parentFolder = System.IO.Directory.GetParent(folderPath).Name;

            //Fix the files 
            if (!string.Equals(parentFolder, "fixfiles", StringComparison.OrdinalIgnoreCase))
            {
                FixFile(folderPath);
            }

            //Parse the xml and create excel file
            GetData(folderPath);

            return ErrorFileList;
        }

        #region Create Excel

        private void GetData(string folderPath)
        {
            var xmlFiles = System.IO.Directory.GetFiles(folderPath + "\\fixFiles\\");
            List<WorkflowData> workflowList = new List<WorkflowData>();

            foreach (string file in xmlFiles)
            {
                GeFileData(file, workflowList);
            }

            WriteWorkFlowFile(workflowList, folderPath);
        }


        private string GeFileData(string filePath, List<WorkflowData> workflowList)
        {
            string returnValue = string.Empty;

            try
            {
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
                    DateTime startTime = DateTime.Today;
                    row.DaysLastModified = (startTime - modifiedDate.Date).Days;

                    workflowList.Add(row);
                }
            }
            catch (Exception ex)
            {
                ErrorFileList.Add(filePath + " - " + ex.Message);
            }
            
            return returnValue;
        }


        private void WriteWorkFlowFile(List<WorkflowData> listData, string folderPath)
        {
            System.IO.StreamWriter writer = new System.IO.StreamWriter(folderPath + "\\" + "workflow.csv");
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

        #endregion

        #region Fix Files


        private void FixFile(string folderPath)
        {

            var xmlFiles = System.IO.Directory.GetFiles(folderPath);

            foreach (string file in xmlFiles)
            {
                FixXml(file, folderPath);
            }

        }

        private void FixXml(string filePath, string folderPath)
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

        # endregion

        #region Utility

        public DateTime FromUnixTime(string unixTimeString)
        {

            long unixTime = long.Parse(unixTimeString);

            var epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            return epoch.AddSeconds(unixTime);
        }

        #endregion

    }
}

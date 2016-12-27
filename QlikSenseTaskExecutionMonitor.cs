using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Configuration;
using System.IO;
using System.Net.Mail;
using System.Net;

namespace QlikSenseTaskExecutionMonitor
{
    public partial class QlikSenseTaskExecutionMonitor : ServiceBase
    {
        private static int vPollingInterval;
        private static string vLogfilePath;
        private static string vLogfileName;
        private static string vSmtpServer;
        private static string vSenderAddress;
        private static string vSenderDisplayName;
        private static string vEmailSubject;
        private static string vRecipientAddresses;
        private static DateTime vLastCheck;
        private static long vLastFileSize;

        private static StreamWriter myLogWriter;
        private static List<string> checkedTasks;

        public static string AssemblyDirectory
        {
            get
            {
                string codeBase = System.Reflection.Assembly.GetExecutingAssembly().Location;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
        }

        public QlikSenseTaskExecutionMonitor()
        {
            InitializeComponent();
        }

        private struct LogfileEntry
        {
            public string cSequence;
            public string cTimestamp;
            public string cLevel;
            public string cHostname;
            public string cLogger;
            public string cThread;
            public string cId;
            public string cServiceUser;
            public string cMessage;
            public string cException;
            public string cStackTrace;
            public string cProxySessionId;
            public string cTaskId;
            public string cTaskName;
            public string cAppId;
            public string cAppName;
            public string cExecutionId;
            public string cExecutingNodeId;
            public string cStatus;
            public string cStartTime;
            public string cStopTime;
            public string cDuration;
            public string cFailureReason;
            public string cId2;
        }

        protected override void OnStart(string[] args)
        {
            myLogWriter = new StreamWriter(Path.Combine(AssemblyDirectory,"QlikSenseTaskExecutionMonitor.log"),true,Encoding.Default);
            myLogWriter.AutoFlush = true;

            myLogWriter.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + ":\tStarting Up Monitoring Service");

            //myLogWriter.WriteLine("CurrentDirectory: " + AssemblyDirectory);

            vLastFileSize = 0;

            if (!ReadConfig())
            {
                ExitCode = 1;
                this.Stop();
            }

            checkedTasks = new List<string>();

            System.Timers.Timer myTimer = new System.Timers.Timer();
            myTimer.Interval = vPollingInterval * 1000; // seconds to milliseconds
            myTimer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnTimer);
            myTimer.Start();
        }

        protected override void OnStop()
        {
            myLogWriter.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + ":\tShutting down Service");
            myLogWriter.Close();
        }

        private bool ReadConfig()
        {
            bool success = true;
            success = int.TryParse(ConfigurationManager.AppSettings["pollingInterval"],out vPollingInterval);

            if (success){
                vLogfilePath = ConfigurationManager.AppSettings["logfilePath"];
                vLogfileName = ConfigurationManager.AppSettings["logfileName"];
                vSmtpServer = ConfigurationManager.AppSettings["smtpServer"];
                vSenderAddress = ConfigurationManager.AppSettings["senderAddress"];
                vSenderDisplayName = ConfigurationManager.AppSettings["senderDisplayName"];
                vRecipientAddresses = ConfigurationManager.AppSettings["recipientAddresses"];
                vEmailSubject = ConfigurationManager.AppSettings["emailSubject"];
            }

            success = (vLogfilePath != null) && (vLogfileName != null) && (vSmtpServer != null) && (vSenderAddress != null) && (vRecipientAddresses != null) && (vEmailSubject != null) && (vSenderDisplayName != null);

            return success;
        }

        public void OnTimer(object sender, System.Timers.ElapsedEventArgs args)
        {
            ScanTaskExecutionLogfile();
        }

        private void ScanTaskExecutionLogfile()
        {
            FileInfo myLogfile;

            try
            {
                myLogfile = new FileInfo(Path.Combine(vLogfilePath, vLogfileName));
                myLogfile.Refresh();
            }
            catch (Exception ex)
            {
                myLogWriter.WriteLine(ex.Message);
                return;
            }

            if (myLogfile.Length == vLastFileSize)
            {
                return;
            }

            myLogWriter.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + ":\tCurrent file size: " + myLogfile.Length + " bytes\tPreviously checked file size: " + vLastFileSize + " bytes");
            //myLogWriter.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + ":\tScanning Task Execution Logfiles ...");


            vLastFileSize = myLogfile.Length;

            string[] separators = { "\r\n" };
            string[] separators2 = {"\t"};

            myLogWriter.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + ":\tAnalyzing Logfile: " + Path.Combine(vLogfilePath, vLogfileName));

            string inFile = null;

            try
            {
                using (var fileStream = new FileStream(Path.Combine(vLogfilePath, vLogfileName),FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var logfileReader = new StreamReader(fileStream, Encoding.Default))
                {
                    inFile = logfileReader.ReadToEnd();
                }
            }
            catch (Exception ex) 
            {
                myLogWriter.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + ":\t" + ex.Message);
                return;
            }


            string[] inLines = inFile.Split(separators, StringSplitOptions.RemoveEmptyEntries);

            for (int i = 1; i < inLines.Count(); i++)
            {
                string[] inCol = inLines[i].Split(separators2,StringSplitOptions.None);

                LogfileEntry myLogfileEntry = new LogfileEntry();

                myLogfileEntry.cSequence = inCol[0];
                myLogfileEntry.cTimestamp = inCol[1];
                myLogfileEntry.cLevel = inCol[2];
                myLogfileEntry.cHostname = inCol[3];
                myLogfileEntry.cLogger = inCol[4];
                myLogfileEntry.cThread = inCol[5];
                myLogfileEntry.cId = inCol[6];
                myLogfileEntry.cServiceUser = inCol[7];
                myLogfileEntry.cMessage = inCol[8];
                myLogfileEntry.cException = inCol[9];
                myLogfileEntry.cStackTrace = inCol[10];
                myLogfileEntry.cProxySessionId = inCol[11];
                myLogfileEntry.cTaskId = inCol[12];
                myLogfileEntry.cTaskName = inCol[13];
                myLogfileEntry.cAppId = inCol[14];
                myLogfileEntry.cAppName = inCol[15];
                myLogfileEntry.cExecutionId = inCol[16];
                myLogfileEntry.cExecutingNodeId = inCol[17];
                myLogfileEntry.cStatus = inCol[18];
                myLogfileEntry.cStartTime = inCol[19];
                myLogfileEntry.cStopTime = inCol[20];
                myLogfileEntry.cDuration = inCol[21];
                myLogfileEntry.cFailureReason = inCol[22];
                myLogfileEntry.cId2 = inCol[23];

                string taskKey = myLogfileEntry.cSequence + "_" + myLogfileEntry.cTimestamp;

                if (myLogfileEntry.cStatus == "FinishedFail" && !(checkedTasks.Contains(taskKey))) {
                    SendEmail(myLogfileEntry);
                    myLogWriter.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + ":\t\tTASK FAILED:\t" + myLogfileEntry.cTaskName + ", APP: " + myLogfileEntry.cAppName + " ... sending Email to " + vRecipientAddresses);
                    checkedTasks.Add(taskKey);
                }

            }

        }

        private void SendEmail(LogfileEntry myLogfileEntry)
        {
            MailMessage myMailMessage = new MailMessage();
            
            string[] separators = {";"};
            string[] recipients = vRecipientAddresses.Split(separators, StringSplitOptions.RemoveEmptyEntries);

            foreach(string recipient in recipients)
                myMailMessage.To.Add(new MailAddress(recipient));

            myMailMessage.Priority = MailPriority.High;
            myMailMessage.From = new MailAddress(vSenderAddress, vSenderDisplayName);
            myMailMessage.Subject = vEmailSubject;
            myMailMessage.Body = "The task '" + myLogfileEntry.cTaskName + "' returned with status 'FinishedFail'";
            myMailMessage.Body += "\r\n\r\nApp Name: " + myLogfileEntry.cAppName;
            myMailMessage.Body += "\r\nExecution time: " + myLogfileEntry.cStartTime;
            myMailMessage.Body += "\r\n\r\nPlease visit https://" + myLogfileEntry.cHostname + "/qmc/tasks to see more details";

            SmtpClient mySmtpClient = new SmtpClient(vSmtpServer);
            mySmtpClient.Send(myMailMessage);

        }
    }
}

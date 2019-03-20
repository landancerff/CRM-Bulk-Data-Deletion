using Boomi.CRM.BulkDataDeletionApp.Controller;
using Microsoft.Xrm.Sdk;
using System;
using System.Diagnostics;

namespace Boomi.CRM.BulkDataDeletionApp.Program
{
    class MainProgram
    {
        public static void Main(string[] args)
        {
            //string environmentSpace = args[0]; 
            string environmentSpace = "extdev3";
            string error = string.Empty;            
            var ctl = new MainController();
            var date = DateTime.Now.ToString("ddMMyyyy_HHmmss"); 

            Trace.AutoFlush = true;
            Trace.Listeners.Clear();
            Trace.Listeners.Add(new ConsoleTraceListener());
            Trace.Listeners.Add(new TextWriterTraceListener($"JobLogs_" + date + ".txt"));
            Trace.Indent();

            var successfulProgram = ctl.ProcessBulkRecordDeletion(environmentSpace, out error);
            Trace.Unindent();

            if (!successfulProgram)
            {
                Trace.WriteLine($"Fatal Error: {error}");                
            }
        }
    }
}

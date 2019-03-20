using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.ServiceModel;
using Boomi.CRM.BulkDataDeletionApp.Connection;
using System.Diagnostics;
using Boomi.CRM.BulkDataDeletionApp.Service;

namespace Boomi.CRM.BulkDataDeletionApp.Controller
{
        public class MainController
    {
        public  bool ProcessBulkRecordDeletion(string environmentSpace, out string errorMsg)
        {
            errorMsg = string.Empty;
            IOrganizationService service = default(IOrganizationService);
            try
            {               
                var repAct = new RepositorService();

                //make CRM Connection
                var crmCnct = new CRMConnection();
                var connected = crmCnct.CreateCRMConnection(out service, environmentSpace, out errorMsg);

                if (!connected)
                {
                    errorMsg = $"No CRM connection established. Please review the URL and Credentials.";
                    return false;
                }              
             
                Trace.WriteLine("Program started; Record deletion initiated.");

                //Get Config data list
                var cnfigServ = new ConfigurationService();
                List<KeyValuePair<string, int>> config = cnfigServ.GetConfigValues(out errorMsg);

                if (config == null || config == default(List<KeyValuePair<string, int>>))
                {
                    errorMsg = $"Configuration list has returned no data. Program can  not continue.";
                    return false;
                }

                //get case types list
                var crmInput = new CRMService();
                List<KeyValuePair<Guid, string>> caseTypes = crmInput.GetCaseTypeList(service, out errorMsg);

                if (caseTypes == null || caseTypes == default(List<KeyValuePair<Guid, string>>))
                {
                    errorMsg = $"Case type list comprised of CRM case types contains no data.";
                    return false;
                }

                // compile config with CRM case type ID's
                List<KeyValuePair<Guid, int>> compiledCaseTypeList = crmInput.MergeCRMCaseTypesWithConfig(caseTypes, config, out errorMsg);

                if (compiledCaseTypeList == null || compiledCaseTypeList == default(List<KeyValuePair<Guid, int>>))
                {
                    errorMsg = $"The comprised list of case type ID's and configuration value amounts contains no data.";
                    return false;
                }

                //get case count values from config, fetch that many cases of x case type
                EntityCollection protectedCases = crmInput.GetProtectedCases(service, compiledCaseTypeList, out errorMsg);
                
                if (protectedCases == null || protectedCases == default(EntityCollection))
                {
                    errorMsg = $"The collection of protected cases has returned containing no data.";
                    return false;
                }

               // bulk update protected cases
                var updateResponse = repAct.BulkUpdateProtectedCases(service, protectedCases, out errorMsg);

                if (!updateResponse)
                {
                    errorMsg = $"There was an error whilst updating all protected cases.";
                    return false;
                }

                //bulk delete unflagged cases
                var recordsDeleted = crmInput.CreateBulkCaseDeletionJob(service, out errorMsg);

                if (!recordsDeleted)
                {
                    errorMsg = $"CRM Bulk Delete for cases message was unsuccessfully sent to CRM.";
                    return false;
                }

                //bulk delete completed system jobs
                var systemJobsComplete = crmInput.CreateSystemJobDeletionJob(service, out errorMsg);

                if (!systemJobsComplete)
                {
                    errorMsg = $"CRM Bulk Delete system jobs was unsuccessfully sent to CRM.";
                    return false;
                }
                    
                Trace.WriteLine("End");
                return true;
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                errorMsg = e.Message;
                Trace.WriteLine(e.Message);
                return false;
            }
        }    
    }
}

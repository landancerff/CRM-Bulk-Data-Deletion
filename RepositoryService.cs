using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using SHG.CD.Plugins.Common.BusinessObjects;
using System.Configuration;
using System.Diagnostics;
using System.ServiceModel;

namespace Boomi.CRM.BulkDataDeletionApp.Service
{
    public class RepositorService
    {
        public bool BulkUpdateProtectedCases(IOrganizationService service, EntityCollection protectedCases, out string errorMsg)
        {
            if (protectedCases == null && protectedCases == default(EntityCollection))
            {
                errorMsg = "The 'Protected Cases' EntityCollection contains no data.";
                return false;
            }
            int count = 0;
            int configLimit = int.Parse(ConfigurationManager.AppSettings["CaseUpdateLimit"]);
            int updateLimit = configLimit > protectedCases.Entities.Count ? protectedCases.Entities.Count : configLimit;

            try
            {
                var multipleRequest = new ExecuteMultipleRequest()
                {
                    Settings = new ExecuteMultipleSettings()
                    {
                        ContinueOnError = false,
                        ReturnResponses = true
                    },
                    Requests = new OrganizationRequestCollection()
                };
                
                foreach (Entity entity in protectedCases.Entities)
                {
                    Entity incident = new Entity { LogicalName = IncidentEntity.LogicalName, Id = entity.Id };
                    incident.Attributes[IncidentEntity.ProtectedCase] = true;
                    UpdateRequest updateRequest = new UpdateRequest { Target = incident };
                    multipleRequest.Requests.Add(updateRequest);

                    if (count == updateLimit)
                    {
                        service.Execute(multipleRequest);
                        Trace.WriteLine($"{count} Records have been updated.");
                        multipleRequest.Requests.Clear();
                        count = 0;
                    } 
                    count++;
                     Trace.WriteLine($"{count} Case: {entity.Attributes[IncidentEntity.CaseNumber]} has been updated.");
                }                
                errorMsg = string.Empty;
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

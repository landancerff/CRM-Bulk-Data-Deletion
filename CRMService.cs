using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using SHG.CD.Plugins.Common.BusinessObjects;
using System.ServiceModel;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;
using System.Diagnostics;
using Boomi.CRM.BulkDataDeletionApp.Repository;
using System.Configuration;

namespace Boomi.CRM.BulkDataDeletionApp.Service
{
    public class CRMService
    {
        public List<KeyValuePair<Guid, string>> GetCaseTypeList(IOrganizationService service, out string errorMsg)
        {
            int caseCount = 0;
            var query = new CRMRepository();
            try
            {
                List<KeyValuePair<Guid, string>> typeStorage = new List<KeyValuePair<Guid, string>>();
                EntityCollection _retrievedTypes = query.RetrieveCRMCaseTypes(service);

                //Trace.WriteLine("\n The following are returned case types from the current CRM environment");
                foreach (Entity e in _retrievedTypes.Entities)
                {
                    //add individual case types to a list
                    string caseName = e.GetAttributeValue<string>(CaseTypeEntity.Name).Replace(" ", string.Empty).ToLower();
                    Guid caseId = e.GetAttributeValue<Guid>(CaseTypeEntity.ID);

                    typeStorage.Add(new KeyValuePair<Guid, string>(caseId, caseName));

                    caseCount++;
                    //Trace.WriteLine($"{caseCount} CRM Case Type: {caseName}, ID: {caseId}");
                }
                errorMsg = null;
                return typeStorage;
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                errorMsg = e.Message;
                Trace.WriteLine(e.Message);
                return null;
            }
        }

        public List<KeyValuePair<Guid, int>> MergeCRMCaseTypesWithConfig(List<KeyValuePair<Guid, string>> crmTypes, List<KeyValuePair<string, int>> config, out string errorMsg)
        {
            //add CRM Type ID to list only where the case type is in the config. 
            List<KeyValuePair<Guid, int>> compiledTypeList = new List<KeyValuePair<Guid, int>>();

            // Check for null in either list  
            if (crmTypes.Select(r => r.Value) == null || config.Select(c => c.Key) == null || config == null || crmTypes == null)
            {
                errorMsg = "Configuration list or CRM Case types list contains no data.";
                Trace.WriteLine("\n ERROR: CRM Case Types list or the Config List do not contain case names. OR either list contains no data. \n");
                return null;
            }
            try
            {
                Trace.WriteLine("\n Creating list of ID's and amounts for case fetch");
                foreach (KeyValuePair<string, int> kvp in config)
                {
                    {
                        var typeCRMKVP = crmTypes.FirstOrDefault(c => c.Value == kvp.Key);

                        if (typeCRMKVP.Key != null)
                        {
                            compiledTypeList.Add(new KeyValuePair<Guid, int>(typeCRMKVP.Key, kvp.Value));
                            Trace.WriteLine($"ID: {typeCRMKVP.Key}, Case Type Name: {kvp.Key}, Amount Protected: {kvp.Value}");
                        }
                    }
                }
                errorMsg = null;
                return compiledTypeList;
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                errorMsg = e.Message;
                Trace.WriteLine(e.Message);
                return null;
            }
        }

        public EntityCollection GetProtectedCases(IOrganizationService service, List<KeyValuePair<Guid, int>> compiledCaseTypeList, out string errorMsg)
        {
            int count = 0;
            EntityCollection protectedCases = new EntityCollection();

            try
            {
                var query = new CRMRepository();
                var input = new CRMService();
                var fetchLimit = int.Parse(ConfigurationManager.AppSettings["CaseFetchLimit"]);

                foreach (KeyValuePair<Guid, int> kpv in compiledCaseTypeList)
                {
                    if (kpv.Value > fetchLimit)
                    {
                        EntityCollection protectedCasesFromPaging = input.GetProtectedCasesWithPaging(service, kpv.Key, kpv.Value, out errorMsg);

                        Trace.WriteLine("\n Retreived cases from paging:");
                        foreach (Entity e in protectedCasesFromPaging.Entities)
                        {
                            protectedCases.Entities.Add(e);
                           // EntityReference returnedCaseType = (EntityReference)e.Attributes[IncidentEntity.CaseType];
                            //Trace.WriteLine($"{count} Case ID: {e.Attributes[IncidentEntity.CaseNumber]}, Case Type: {returnedCaseType.Name}");
                            count++;
                        }
                        count = 0;
                    }
                    else
                    {
                        EntityCollection returnedCases = query.GetCasesForType(service, kpv.Key, kpv.Value);
                        //Trace.WriteLine("\n Retreived CRM cases:");
                        if(returnedCases != null && returnedCases.Entities.Count > 0)
                        {
                            foreach (Entity e in returnedCases.Entities)
                            {
                                protectedCases.Entities.Add(e);
                                //EntityReference returnedCaseType = (EntityReference)e.Attributes[IncidentEntity.CaseType];
                                //Trace.WriteLine($"{count} Case ID: {e.Attributes[IncidentEntity.CaseNumber]}, Case Type: {returnedCaseType.Name}");
                                count++;
                            }                            
                        }                                                
                        count = 0;
                    }
                }
                if (protectedCases != null && protectedCases != default(EntityCollection))
                {
                    errorMsg = null;
                    return protectedCases;
                }
                errorMsg = null;
                return null;
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                errorMsg = e.Message;
                Trace.WriteLine(e.Message);
                return null;
            }
        }

        public EntityCollection GetProtectedCasesWithPaging(IOrganizationService service, Guid key, int value, out string errorMsg)
        {
            EntityCollection protectedCasesFromPaging = new EntityCollection();
            try
            {
                if (key == null && key == default(Guid) && value != 0 && value != default(int))
                {
                    errorMsg = "Key or Value from Config is missing";
                    return null;
                }
                var query = new CRMRepository();
                protectedCasesFromPaging = query.RecordFetchPagingQuery(service, value, key);
                if (protectedCasesFromPaging != null)
                {
                    errorMsg = null;
                    return protectedCasesFromPaging;
                }
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                errorMsg = e.Message;
                Trace.WriteLine(e.Message);
                return null;
            }
            errorMsg = null;
            return protectedCasesFromPaging;
        }

        public bool CreateBulkCaseDeletionJob(IOrganizationService service, out string errorMsg)
        {
            try
            {
                ConditionExpression conditionExp = new ConditionExpression(IncidentEntity.ProtectedCase, ConditionOperator.NotEqual, true);
                FilterExpression filterExp = new FilterExpression();
                filterExp.AddCondition(conditionExp);
                BulkDeleteRequest request = new BulkDeleteRequest
                {
                    JobName = "Delete Unprotected Records - Bulk Case Deletion",
                    ToRecipients = new Guid[] { },                  
                    CCRecipients = new Guid[] { },
                    RunNow = true,
                    RecurrencePattern = string.Empty,
                    QuerySet = new QueryExpression[]
            {
             new QueryExpression { EntityName = IncidentEntity.LogicalName, Criteria = filterExp}
            }
                };
                BulkDeleteResponse response = (BulkDeleteResponse)service.Execute(request);
                Trace.WriteLine($"Bulk Deletion Job response: {response.ToString()}");
                errorMsg = null;
                return true;
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                errorMsg = e.Message;
                Trace.WriteLine(e.Message);
                return false;
            }
        }

        public bool CreateSystemJobDeletionJob(IOrganizationService service, out string errorMsg)
        {
            try
            {
                ConditionExpression conEx = new ConditionExpression("statecode", ConditionOperator.Equal, 3);
                FilterExpression fiEx = new FilterExpression();
                fiEx.AddCondition(conEx);

                BulkDeleteRequest request = new BulkDeleteRequest
                {
                    JobName = "Bulk Delete Completed System Jobs",                   
                    CCRecipients = new Guid[] { },
                    ToRecipients = new Guid[] { },
                    RecurrencePattern = string.Empty,
                    QuerySet = new QueryExpression[]
            {
            new QueryExpression { EntityName = "asyncoperation", Criteria = fiEx}
            }
                };
                BulkDeleteResponse response = (BulkDeleteResponse)service.Execute(request);
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

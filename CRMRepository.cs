using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using SHG.CD.Plugins.Common.BusinessObjects;
using System;

namespace Boomi.CRM.BulkDataDeletionApp.Repository
{
    public class CRMRepository
    {        
        public  EntityCollection RetrieveCRMCaseTypes(IOrganizationService service)
        {
            string caseTypeFetch = string.Empty;

            caseTypeFetch = @"<fetch distinct='false' mapping='logical' output-format='xml-platform' version='1.0'>
                                  <entity name = 'gcs_casetype'>
                                   <attribute name = 'gcs_casetypeid'/>
                                    <attribute name = 'gcs_name'/>
                                     <attribute name = 'createdon'/>
                                      <order descending = 'false' attribute = 'gcs_name'/>
                                         </entity>
                                         </fetch>";

            FetchExpression f = new FetchExpression(caseTypeFetch);
            if (f != null && f != default(FetchExpression))
            {
                EntityCollection activeTypes = service.RetrieveMultiple(f);
                if (activeTypes.Entities.Count > 0)
                {
                    return activeTypes;
                }
                return null;
            }
            return null;
        }

        public  EntityCollection GetCasesForType(IOrganizationService service, Guid categoryId, int recordFetchCount)
        {
            string caseFetch = string.Empty;
            caseFetch = $@"<fetch top='{recordFetchCount}' distinct='true'>
                              <entity name='incident'>
                                <attribute name='incidentid' />
                                <attribute name='ticketnumber' />
                                <attribute name='gcs_casetypes' />
                                <attribute name='statecode'/>
                                <attribute name='shg_casesavefield' />
                                <filter type='and'>
                                  <condition attribute='gcs_casetypes' operator='eq' value='{categoryId}'/>
                                    <condition attribute='statecode' operator='eq' value='{0}'/>                               
                                </filter>
                              </entity>
                            </fetch>";

            FetchExpression f = new FetchExpression(caseFetch);

            EntityCollection cases = service.RetrieveMultiple(f);
            if (cases.Entities.Count > 0)
            {
                return cases;
            }
           return null;            
        }

        public  EntityCollection RecordFetchPagingQuery(IOrganizationService service, int recordTotal, Guid caseType)
        {        
            int queryCount = 60;
            int pageNumber = 1;
            int recordCount = 0;
            EntityCollection retrievedRecords = new EntityCollection();

            ConditionExpression pagecondition = new ConditionExpression();
            pagecondition.AttributeName = IncidentEntity.CaseType;
            pagecondition.Operator = ConditionOperator.Equal;
            pagecondition.Values.Add(caseType);

            OrderExpression order = new OrderExpression();
            order.AttributeName = IncidentEntity.CreatedOn;
            order.OrderType = OrderType.Ascending;

            QueryExpression pagequery = new QueryExpression();

            pagequery.EntityName = IncidentEntity.LogicalName;
            pagequery.Criteria.AddCondition(pagecondition);
            pagequery.Orders.Add(order);
            pagequery.ColumnSet.AddColumns(IncidentEntity.ID, IncidentEntity.CaseType, IncidentEntity.CaseNumber);
            pagequery.PageInfo = new PagingInfo();
            pagequery.PageInfo.Count = queryCount;
            pagequery.PageInfo.PageNumber = pageNumber;
            pagequery.PageInfo.PagingCookie = null;

            while (true)
            {
                EntityCollection returnedCases = service.RetrieveMultiple(pagequery);

                if(returnedCases.Entities.Count == 0)
                {
                    break;
                }
                if (returnedCases.Entities != null)
                {
                    foreach (Entity e in returnedCases.Entities)
                    {                        
                        retrievedRecords.Entities.Add(e);
                        recordCount++;
                    }
                }              
                if (retrievedRecords.Entities.Count < recordTotal)
                {
                    int difference = Math.Abs(recordTotal - recordCount);
                   
                    if(difference <= queryCount)
                    {
                        pagequery.PageInfo.Count = difference;
                    }                   
                    pagequery.PageInfo.PageNumber++;
                    pagequery.PageInfo.PagingCookie = returnedCases.PagingCookie;
                }
                else
                {                   
                   break;
                }           
            }
            return retrievedRecords;
        }
    }
}

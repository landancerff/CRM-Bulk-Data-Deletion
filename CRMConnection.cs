using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System;
using System.Configuration;
using System.Diagnostics;
using System.Net;
using System.ServiceModel;
using System.ServiceModel.Description;

namespace Boomi.CRM.BulkDataDeletionApp.Connection
{
    public class CRMConnection
    {
        public  bool CreateCRMConnection(out IOrganizationService service, string environmentSpace, out string error)
        {
            service = default(IOrganizationService);       
            try
            {
                ClientCredentials clientCredentials = new ClientCredentials();
                clientCredentials.UserName.UserName = ConfigurationManager.AppSettings["AdminEmail"];
                clientCredentials.UserName.Password = ConfigurationManager.AppSettings["AdminPassword"];

                Trace.WriteLine("CRM Credentials set.");
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                service = (IOrganizationService)new OrganizationServiceProxy(
                    new Uri("https://shgl-" + environmentSpace + ".api.crm4.dynamics.com/XRMServices/2011/Organization.svc"),
                    null, clientCredentials, null);

                Trace.WriteLine("Successful CRM connection established.");
                error = null;
                return true;
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                error = e.Message;
                Trace.WriteLine(e.Message);              
            }
            return false;
        }
    }
}




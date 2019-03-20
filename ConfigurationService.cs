using Microsoft.Xrm.Sdk;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.ServiceModel;

namespace Boomi.CRM.BulkDataDeletionApp.Service
{
    class ConfigurationService
    {
        public List<KeyValuePair<string, int>> GetConfigValues(out string errorMsg)
        {
            try
            {
                List<KeyValuePair<string, int>> configStorage = null;
                configStorage = new List<KeyValuePair<string, int>>();

                var sections = ConfigurationManager.GetSection("MajorCommands") as System.Collections.Hashtable;

                var kvpList = sections.Cast<DictionaryEntry>().ToDictionary(kvp => (string)kvp.Key, kvp => int.Parse(kvp.Value.ToString()));
                Trace.WriteLine("The following case types have config, these values will be used to save the set amount of cases of each type.");
                foreach (DictionaryEntry kvp in sections)
                {
                    if (kvp.Value != null && kvp.Key != null)
                    {
                        configStorage.Add(new KeyValuePair<string, int>(kvp.Key.ToString(), int.Parse(kvp.Value.ToString())));
                        Trace.WriteLine($"Case Type: {kvp.Key} Count: {kvp.Value}");
                    }
                }
                errorMsg = null;
                return configStorage;
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                errorMsg = e.Message;
                Trace.WriteLine(e.Message);
                return null;
            }
        }
    }
}

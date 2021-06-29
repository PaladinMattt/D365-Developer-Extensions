using D365DeveloperExtensions.Core.Connection;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Tooling.Connector;

namespace D365DeveloperExtensions.Core.Crm
{
    public static class Connection
    {
        public static string RetrieveOrganizationId(CrmConnect service)
        {
            var response = (WhoAmIResponse)service.Service.Execute(new WhoAmIRequest());

            return response.OrganizationId.ToString();
        }
    }
}
using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Extensions;

namespace POAVE.Plugins
{
    public class Retrieve : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = serviceProvider.Get<IPluginExecutionContext>();
            var service = serviceProvider.GetOrganizationService(context.UserId);

            var target = (EntityReference)context.InputParameters["Target"];

            var query = new QueryByAttribute("principalobjectaccess")
            {
                ColumnSet = new ColumnSet(true)
            };
            query.AddAttributeValue("principalobjectaccessid", target.Id);
            var poaRecords = service.RetrieveMultiple(query).Entities.ToList();
            var poaResults = poaRecords.Select(p => new Entity("ab_poa", p.Id)
            {
                ["ab_poaid"] = p.Id,
                ["ab_objectid"] = p.GetAttributeValue<Guid>("objectid").ToString(),
                ["ab_objecttypecode"] = p.GetAttributeValue<string>("objecttypecode"),
                ["ab_principalid"] = p.GetAttributeValue<Guid>("principalid").ToString(),
                ["ab_principaltypecode"] = p.GetAttributeValue<string>("principaltypecode"),
                ["ab_read"] = (p.GetAttributeValue<int>("accessrightsmask") & (int)AccessRights.ReadAccess) != 0,
                ["ab_write"] = (p.GetAttributeValue<int>("accessrightsmask") & (int)AccessRights.WriteAccess) != 0,
                ["ab_delete"] = (p.GetAttributeValue<int>("accessrightsmask") & (int)AccessRights.DeleteAccess) != 0,
                ["ab_append"] = (p.GetAttributeValue<int>("accessrightsmask") & (int)AccessRights.AppendAccess) != 0,
                ["ab_appendto"] = (p.GetAttributeValue<int>("accessrightsmask") & (int)AccessRights.AppendToAccess) != 0,
                ["ab_assign"] = (p.GetAttributeValue<int>("accessrightsmask") & (int)AccessRights.AssignAccess) != 0,
                ["ab_share"] = (p.GetAttributeValue<int>("accessrightsmask") & (int)AccessRights.ShareAccess) != 0
            }).ToList();

            context.OutputParameters["BusinessEntity"] = poaResults.First();
        }
    }
}

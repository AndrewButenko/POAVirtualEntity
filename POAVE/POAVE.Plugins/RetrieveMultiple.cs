using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Extensions;

namespace POAVE.Plugins
{
    public class RetrieveMultiple : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = serviceProvider.Get<IPluginExecutionContext>();
            var service = serviceProvider.GetOrganizationService(context.UserId);

            var sourceQuery = (QueryExpression) context.InputParameters["Query"];

            var query = new QueryExpression("principalobjectaccess")
            {
                ColumnSet = new ColumnSet(true),
                Criteria = ConvertFilter(sourceQuery.Criteria),
                PageInfo = sourceQuery.PageInfo
            };

            query.Orders.AddRange(sourceQuery.Orders.Select(o => ConvertOrder(o)).Where(o => o != null));

            var poaRecords = service.RetrieveMultiple(query);

            var poaResults = poaRecords.Entities.Select(p => new Entity("ab_poa", p.Id)
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

            var result = new EntityCollection()
            {
                EntityName = context.PrimaryEntityName,
                TotalRecordCount = poaResults.Count,
                MoreRecords = poaRecords.MoreRecords,
                PagingCookie = poaRecords.PagingCookie
            };
            result.Entities.AddRange(poaResults);

            context.OutputParameters["BusinessEntityCollection"] = result;
        }

        private FilterExpression ConvertFilter(FilterExpression sourceFilter)
        {
            var result = new FilterExpression(sourceFilter.FilterOperator);

            result.Conditions.AddRange(sourceFilter.Conditions.Select(c => ConvertCondition(c)));
            result.Filters.AddRange(sourceFilter.Filters.Select(f => ConvertFilter(f)));

            return result;
        }

        private ConditionExpression ConvertCondition(ConditionExpression sourceCondition)
        {
            var result = new ConditionExpression()
            {
                Operator = sourceCondition.Operator
            };

            switch (sourceCondition.AttributeName)
            {
                case "ab_objectid":
                    result.AttributeName = "objectid";
                    result.Values.AddRange(sourceCondition.Values.Select(v => Guid.Parse(v.ToString())));
                    break;
                case "ab_objecttypecode":
                    result.AttributeName = "objecttypecode";
                    result.Values.AddRange(sourceCondition.Values);
                    break;
                case "ab_principalid":
                    result.AttributeName = "principalid";
                    result.Values.AddRange(sourceCondition.Values.Select(v => Guid.Parse(v.ToString())));
                    break;
                case "ab_principaltypecode":
                    result.AttributeName = "principaltypecode";
                    result.Values.AddRange(sourceCondition.Values);
                    break;
                default:
                    throw new InvalidPluginExecutionException(
                        $"Attribute {sourceCondition.AttributeName} can't be used in conditions");
            }

            return result;
        }

        private OrderExpression ConvertOrder(OrderExpression sourceOrder)
        {
            switch (sourceOrder.AttributeName)
            {
                case "ab_objectid":
                    return new OrderExpression("objectid", sourceOrder.OrderType);
                case "ab_objecttypecode":
                    return new OrderExpression("objecttypecode", sourceOrder.OrderType);
                case "ab_principalid":
                    return new OrderExpression("principalid", sourceOrder.OrderType);
                case "ab_principaltypecode":
                    return new OrderExpression("principaltypecode", sourceOrder.OrderType);
            }

            return null;
        }
    }
}

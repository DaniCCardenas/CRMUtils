using System;
using System.Activities;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;

namespace DCC.CRM.Utils.GetEntityValue
{
    public class GetEntityValue : CodeActivity
    {
        [RequiredArgument]
        [Input("Entity Name")]
        public InArgument<string> EntityName { get; set; }

        [RequiredArgument]
        [Input("Entity Key")]
        public InArgument<string> EntityKey { get; set; }

        [RequiredArgument]
        [Input("Entity Value")]
        public InArgument<string> EntityValue { get; set; }

        [RequiredArgument]
        [Input("Search Value")]
        public InArgument<string> SearchValue { get; set; }

        [Output("Value")]
        public OutArgument<string> Value { get; set; }


        protected override void Execute(CodeActivityContext executionContext)
        {
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);


            var query = new QueryExpression(EntityName.Get(executionContext)) {TopCount = 1};
            query.ColumnSet.AddColumns(EntityValue.Get(executionContext));
            query.Criteria.AddCondition(EntityKey.Get(executionContext), ConditionOperator.Equal, SearchValue.Get(executionContext));

            var resutl = service.RetrieveMultiple(query);

            if (!resutl.Entities.Any())
            {
                throw new Exception("Value not found");
            }
            if (resutl.Entities.Count > 1)
            {
                throw new Exception("There are more than one result");
            }


            Value.Set(executionContext, resutl.Entities[0].GetAttributeValue<string>(EntityValue.Get(executionContext)));
        }
    }
}

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Linq;

namespace D365_Core_Workflows.WorkflowActivities
{
    public class IsUserInSelectedTeam : CodeActivity
    {
        [RequiredArgument]
        [Input("Tean")]
        [ReferenceTarget("team")]
        public InArgument<EntityReference> Team { get; set; }

        [RequiredArgument]
        [Input("User")]
        [ReferenceTarget("systemuser")]
        public InArgument<EntityReference> User { get; set; }

        [Output("Result")]
        public OutArgument<bool> Result { get; set; }

        #region Service Parameters
        private ITracingService tracingService;
        private IWorkflowContext context;
        private IOrganizationServiceFactory serviceFactory;
        private IOrganizationService service;
        #endregion Service Parameters

        protected override void Execute(CodeActivityContext executionContext)
        {
            #region Initializing Services
            try
            {
                tracingService = executionContext.GetExtension<ITracingService>();
                context = executionContext.GetExtension<IWorkflowContext>();
                serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
                service = serviceFactory.CreateOrganizationService(context.UserId);

                tracingService.Trace("Services are initialized.");
            }
            catch (Exception e)
            {

                throw new InvalidOperationException($"There was an error during initializing the services:  {e.Message}");
            }


            #endregion Initializing Services

            #region Parameters
            tracingService.Trace("Reading Input Parameters");

            var teamRef = this.Team.Get(executionContext);
            var userRef = this.User.Get(executionContext);

            if (teamRef.Id == Guid.Empty || userRef.Id == Guid.Empty)
                throw new InvalidPluginExecutionException("Invalid input parameters! Please contact with your System Administrator.");

            #endregion Parameters

            #region Query Team Members
            tracingService.Trace($"Starting Queryig Team Members for the {teamRef.Id}");

            var query = new QueryExpression()
            {
                EntityName = "teammembership",
                ColumnSet = new ColumnSet(true),
                Criteria = new FilterExpression(LogicalOperator.And)
                {
                    Conditions =
                    {
                        new ConditionExpression("teamid", ConditionOperator.Equal, teamRef.Id),
                    }
                }
            };

            query.LinkEntities.Add(new LinkEntity("teammembership", "systemuser", "systemuserid", "systemuserid", JoinOperator.Inner));
            query.LinkEntities[0].LinkCriteria.AddCondition("systemuserid", ConditionOperator.Equal, userRef.Id);
            query.LinkEntities[0].Columns.AddColumns("firstname", "lastname");
            query.LinkEntities[0].EntityAlias = "systemuser";

            var result = service.RetrieveMultiple(query).Entities;

            tracingService.Trace("Ending Queryig Team Members");
            #endregion Query Team Members

            this.Result.Set(executionContext, result.Any());
        }
    }
}

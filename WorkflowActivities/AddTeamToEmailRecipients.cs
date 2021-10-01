using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Linq;


namespace D365_Core_Workflows.WorkflowActivities
{
    public class AddTeamToEmailRecipients : CodeActivity
    {
        [RequiredArgument]
        [Input("Tean")]
        [ReferenceTarget("team")]
        public InArgument<EntityReference> Team { get; set; }

        [RequiredArgument]
        [Input("Email")]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> Email { get; set; }

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
            var emailRef = this.Email.Get(executionContext);

            if (teamRef.Id == Guid.Empty || emailRef.Id == Guid.Empty)
                throw new InvalidPluginExecutionException("Invalid input parameters! Please contact with your System Administrator.");

            #endregion Parameters

            #region Retrieve Team Members
            tracingService.Trace("Starting Retrieving Team Members...");

            var teamMembers = service.RetrieveMultiple(new QueryExpression("teammembership")
            {
                ColumnSet = new ColumnSet(true),
                Criteria = new FilterExpression(LogicalOperator.And)
                {
                    Conditions =
                        {
                            new ConditionExpression("teamid", ConditionOperator.Equal, teamRef.Id)
                        }
                }
            }).Entities;

            tracingService.Trace("Team Members are Retrieved...");

            if (!teamMembers.Any())
            {
                tracingService.Trace($"The Team with the ID: {teamRef.Id} does not contain any users.");
                return;
            }

            #endregion Retrieve Team Members

            #region Update Email

            EntityCollection recipientsCollection = new EntityCollection();

            foreach (var membership in teamMembers)
            {
                Entity to = new Entity("activityparty");
                to["partyid"] = new EntityReference("systemuser", membership.GetAttributeValue<Guid>("systemuserid"));

                recipientsCollection.Entities.Add(to);
            }

            tracingService.Trace($"Starting Update Email Recipients for the Email with the ID: {emailRef.Id} ...");

            Entity updateEmail = new Entity("email", emailRef.Id);
            updateEmail["to"] = recipientsCollection;
            service.Update(updateEmail);

            tracingService.Trace($"Updated Email Recipients...");
            #endregion Update Email
        }
    }
}

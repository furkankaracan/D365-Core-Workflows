using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace D365_Core_Workflows.WorkflowActivities
{
    public class AddSelectedUserToTeam : CodeActivity
    {
        [RequiredArgument]
        [Input("User")]
        [ReferenceTarget("systemuser")]
        public InArgument<EntityReference> User { get; set; }

        [RequiredArgument]
        [Input("Team")]
        [ReferenceTarget("team")]
        public InArgument<EntityReference> Team { get; set; }

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

            var userRef = this.User.Get(executionContext);
            var teamRef = this.Team.Get(executionContext);

            if (userRef.Id == Guid.Empty || teamRef.Id == Guid.Empty)
                throw new InvalidPluginExecutionException("Invalid input parameters! Please contact with your System Administrator.");

            #endregion Parameters

            #region Add User To Team
            tracingService.Trace("Starting AddMembersTeamRequest");

            AddMembersTeamRequest addMembersTeamRequest = new AddMembersTeamRequest()
            {
                MemberIds = new[] { userRef.Id },
                TeamId = teamRef.Id
            };

            _ = (AddMembersTeamResponse)service.Execute(addMembersTeamRequest);

            tracingService.Trace("Ending AddMembersTeamRequest");
            #endregion Add User To Team
        }
    }
}
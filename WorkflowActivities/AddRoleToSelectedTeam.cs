using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Linq;

namespace D365_Core_Workflows.WorkflowActivities
{
    public class AddRoleToSelectedTeam : CodeActivity
    {
        [RequiredArgument]
        [Input("Tean")]
        [ReferenceTarget("team")]
        public InArgument<EntityReference> Team { get; set; }

        [RequiredArgument]
        [Input("Role")]
        [ReferenceTarget("role")]
        public InArgument<EntityReference> Role { get; set; }

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
            var roleRef = this.Role.Get(executionContext);

            if (teamRef.Id == Guid.Empty || roleRef.Id == Guid.Empty)
                throw new InvalidPluginExecutionException("Invalid input parameters! Please contact with your System Administrator.");

            #endregion Parameters

            #region Retrieve Team
            tracingService.Trace("Retrieving Team ");

            Entity team = service.Retrieve("team", teamRef.Id, new ColumnSet("businessunitid"));

            EntityReference businessUnitRef = team.GetAttributeValue<EntityReference>("businessunitid");

            #endregion Retrieve Team

            #region Retrieve Roles
            tracingService.Trace("Retrieving Team Root Roles ");

            var teamRootRoles = service.RetrieveMultiple(new QueryExpression("role")
            {
                ColumnSet = new ColumnSet("parentrootroleid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("roleid", ConditionOperator.Equal, roleRef.Id)
                    }
                }
            }).Entities;

            tracingService.Trace("Retrieved Team Root Roles ");
            #endregion Retrieve Roles

            if (teamRootRoles.Any())
            {
                Entity teamRole = teamRootRoles[0];
                EntityReference parentRootRoleRef = teamRole.GetAttributeValue<EntityReference>("parentrootroleid");

                #region Retrieve Root Roles
                tracingService.Trace("Retrieving Root Roles ");

                // Retrieve the correct role for that businessunit
                var rootRoles = service.RetrieveMultiple(new QueryExpression("role")
                {
                    ColumnSet = new ColumnSet("roleid"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression("parentrootroleid", ConditionOperator.Equal, parentRootRoleRef.Id),
                            new ConditionExpression("businessunitid", ConditionOperator.Equal, businessUnitRef.Id),
                        }
                    }
                }).Entities;

                tracingService.Trace("Retrieving Root Roles ");

                Entity role = rootRoles[0];

                #endregion Retrieve Root Roles

                tracingService.Trace("Adding Role to the Team");

                service.Associate("team",
                    teamRef.Id,
                    new Relationship("teamroles_association"),
                    new EntityReferenceCollection() { role.ToEntityReference() }
                    );

                tracingService.Trace("Role has added to the Team");
            }
        }
    }
}

using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Linq;

namespace D365_Core_Workflows
{
    public class AddRoleToSelectedUser : CodeActivity
    {
        [RequiredArgument]
        [Input("User")]
        [ReferenceTarget("systemuser")]
        public InArgument<EntityReference> User { get; set; }

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

            var userRef = this.User.Get(executionContext);
            var roleRef = this.Role.Get(executionContext);

            if (userRef.Id == Guid.Empty || roleRef.Id == Guid.Empty)
                throw new InvalidPluginExecutionException("Invalid input parameters! Please contact with your System Administrator.");

            #endregion Parameters

            #region Retrieve User
            tracingService.Trace("Retrieving User ");

            Entity systemUser = service.Retrieve("systemuser", userRef.Id, new ColumnSet("businessunitid"));

            EntityReference businessUnit = systemUser.GetAttributeValue<EntityReference>("businessunitid");

            #endregion Retrieve User

            #region Retrieve Roles
            tracingService.Trace("Retrieving User Root Roles ");

            var userRootRoles = service.RetrieveMultiple(new QueryExpression("role")
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

            tracingService.Trace("Retrieved User Root Roles ");
            #endregion Retrieve Roles

            if (userRootRoles.Any())
            {
                Entity userRole = userRootRoles[0];
                EntityReference parentRootRoleRef = userRole.GetAttributeValue<EntityReference>("parentrootroleid");

                #region Retrieve Root Roles
                tracingService.Trace("Retrieving Root Roles ");

                var rootRoles = service.RetrieveMultiple(new QueryExpression("role")
                {
                    ColumnSet = new ColumnSet("roleid"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression("parentrootroleid", ConditionOperator.Equal, parentRootRoleRef.Id),
                            new ConditionExpression("businessunitid", ConditionOperator.Equal, businessUnit.Id),
                        }
                    }
                }).Entities;

                tracingService.Trace("Retrieving Root Roles ");

                Entity role = rootRoles[0];

                #endregion Retrieve Root Roles

                tracingService.Trace("Adding Role to the User");

                service.Associate("systemuser",
                    userRef.Id,
                    new Relationship("systemuserroles_association"),
                    new EntityReferenceCollection() { role.ToEntityReference() }
                    );

                tracingService.Trace("Role has added to the User");
            }
        }
    }
}

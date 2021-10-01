using System;
using System.Linq;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace D365_Core_Workflows
{
    public class CalculateRollupField : CodeActivity
    {
        [RequiredArgument]
        [Input("RollupFieldName")]
        public InArgument<String> RollupFieldName { get; set; }

        [RequiredArgument]
        [Input("Record URL")]
        public InArgument<String> RecordURL { get; set; }

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

            var fieldName = this.RollupFieldName.Get(executionContext);
            var recordURL = this.RecordURL.Get(executionContext);

            if (string.IsNullOrEmpty(fieldName) || string.IsNullOrEmpty(recordURL))
                throw new InvalidPluginExecutionException("Input parameters cannot be null!");

            #endregion Parameters

            #region Formatting Parameters
            tracingService.Trace("Starting Formatting Parameters");

            string[] urlParts = recordURL.Split('?');
            if (!urlParts.Any())
                throw new InvalidPluginExecutionException("Input URL is incorrect! Please contact with your System Administrator.");

            string[] urlParameters = urlParts[1].Split('&');

            int parentObjTypeCode = Convert.ToInt32(urlParameters[0].Replace("etc=", ""));
            string recordId = urlParameters[1].Replace("id=", "");

            if (string.IsNullOrEmpty(recordId))
                throw new InvalidPluginExecutionException("Input URL is incorrect! Please contact with your System Administrator.");

            tracingService.Trace("Ending Formatting Parameters");
            #endregion Formatting Parameters


            #region Calculate Rollup
            tracingService.Trace("Starting Calculating Rollup");

            string ParentEntityName = GetEntityNameFromTypeCode(parentObjTypeCode);
            CalculateRollupFieldRequest calculateRollupFieldRequest = new CalculateRollupFieldRequest()
            {
                FieldName = fieldName,
                Target = new EntityReference(ParentEntityName, new Guid(recordId))
            };

            _ = (CalculateRollupFieldResponse)service.Execute(calculateRollupFieldRequest);

            tracingService.Trace("Ending Calculating Rollup");
            #endregion Calculate Rollup
        }

        public string GetEntityNameFromTypeCode(int objectTypeCode)
        {
            MetadataFilterExpression entityFilter = new MetadataFilterExpression(LogicalOperator.And);
            entityFilter.Conditions.Add(new MetadataConditionExpression("ObjectTypeCode", MetadataConditionOperator.Equals, objectTypeCode));
            EntityQueryExpression entityQueryExpression = new EntityQueryExpression()
            {
                Criteria = entityFilter
            };
            RetrieveMetadataChangesRequest retrieveMetadataChangesRequest = new RetrieveMetadataChangesRequest()
            {
                Query = entityQueryExpression,
                ClientVersionStamp = null
            };
            RetrieveMetadataChangesResponse response = (RetrieveMetadataChangesResponse)service.Execute(retrieveMetadataChangesRequest);

            EntityMetadata entityMetadata = (EntityMetadata)response.EntityMetadata[0];
            return entityMetadata.SchemaName.ToLower();
        }
    }

}

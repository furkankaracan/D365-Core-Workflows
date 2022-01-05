using D365_Core_Workflows.Helpers;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace D365_Core_Workflows.WorkflowActivities
{
    public class SetMultiSelectOptionSet : CodeActivity
    {
        #region WFA Parameters

        [RequiredArgument]
        [Input("Target Record URL")]
        public InArgument<string> TargetRecordUrl { get; set; }

        [RequiredArgument]
        [Input("Attribute Name")]
        public InArgument<string> AttributeName { get; set; }

        [RequiredArgument]
        [Input("Attribute Values")]
        public InArgument<string> AttributeValues { get; set; }

        [Input("Keep Existing Values")]
        [Default("false")]
        public InArgument<Boolean> KeepExistingValues { get; set; }

        #endregion WFA Parameters

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

                throw new InvalidOperationException($"There was an error during initializing the services: {e.Message}");
            }
            #endregion Initializing Services

            EntityReference sourceEntityReference = GetTargeteEntityReference(executionContext);
            string attributeName = GetAttributeName(executionContext);
            OptionSetValueCollection newValues = GetNewAttributeValues(executionContext);
            OptionSetValueCollection existingValues = GetExistingAttributeValues(sourceEntityReference, attributeName, executionContext);

            UpdateRecord(sourceEntityReference, attributeName, newValues, existingValues);
        }

        private EntityReference GetTargeteEntityReference(CodeActivityContext executionContext)
        {
            string sourceRecordUrl = TargetRecordUrl.Get<string>(executionContext) ?? throw new ArgumentNullException("Source URL is empty");
            tracingService.Trace("Source Record URL:'{0}'", sourceRecordUrl);
            return new DynamicUrlParser(sourceRecordUrl).ToEntityReference(service);
        }

        private string GetAttributeName(CodeActivityContext executionContext)
        {
            string attributeName = AttributeName.Get<string>(executionContext) ?? throw new ArgumentNullException("Attribute Name is empty");
            tracingService.Trace("Attribute name:'{0}'", attributeName);
            return attributeName;
        }

        private OptionSetValueCollection GetNewAttributeValues(CodeActivityContext executionContext)

        {
            string attributeValues = AttributeValues.Get<string>(executionContext) ?? throw new ArgumentNullException("Attribute Values is empty");
            tracingService.Trace("Attribute Values:'{0}'", attributeValues);

            if (string.IsNullOrEmpty(attributeValues))
            {
                tracingService.Trace("No values found. Setting attribute to null");
                return new OptionSetValueCollection();
            }

            string[] values = attributeValues.Split(',');

            if (values == null || values.Length == 0)
            {
                tracingService.Trace("No values found in array. Setting attribute to null");
                return new OptionSetValueCollection();
            }

            OptionSetValueCollection optionSetValueCollection = new OptionSetValueCollection();

            foreach (string value in values)
            {
                if (int.TryParse(value, out int intValue))
                {
                    tracingService.Trace("Value '{0}' added correctly", value);
                    optionSetValueCollection.Add(new OptionSetValue(intValue));
                }
                else
                {
                    tracingService.Trace("Value '{0}' couldn't be parsed", value);
                }
            }

            return optionSetValueCollection;
        }

        private OptionSetValueCollection GetExistingAttributeValues(EntityReference targetEntityReference, string attributeName, CodeActivityContext executionContext)
        {
            tracingService.Trace("Retrieving existing values");

            Boolean attributeValues = KeepExistingValues.Get<Boolean>(executionContext);

            if (attributeValues == false)
                return null;

            Entity record = service.Retrieve(targetEntityReference.LogicalName, targetEntityReference.Id, new ColumnSet(new string[] { attributeName }));

            tracingService.Trace("Existing values have been retrieved correctly");

            if (record.Contains(attributeName))
                return record[attributeName] as OptionSetValueCollection;
            else
                return null;
        }

        private void UpdateRecord(EntityReference targetEntityReference, string attributeName, OptionSetValueCollection newValues, OptionSetValueCollection existingValues)
        {
            if (targetEntityReference == null || attributeName == null || newValues == null)
                throw new ArgumentNullException(string.Format("Unexpected null parameters when trying to update record. Record reference '{0}' - attibute name '{1}' - values '{2}'", targetEntityReference, attributeName, newValues));


            Entity targetEntity = new Entity(targetEntityReference.LogicalName, targetEntityReference.Id);
            targetEntity[attributeName] = MergeOptionSetCollections(newValues, existingValues);

            service.Update(targetEntity);

            tracingService.Trace("Multi-select option set attribute '{0}' has been updated correctly for the record type '{1}' with id '{2}'", attributeName, targetEntityReference.LogicalName, targetEntityReference.Id);
        }

        private OptionSetValueCollection MergeOptionSetCollections(OptionSetValueCollection newValues, OptionSetValueCollection existingValues)
        {
            tracingService.Trace("Merging new and exiting multi-select optionset values");

            if (existingValues == null && newValues == null)
                return new OptionSetValueCollection();

            if (existingValues == null)
                return newValues;

            if (newValues == null)
                return existingValues;

            foreach (OptionSetValue newValue in newValues)
            {
                if (!existingValues.Contains(newValue))
                    existingValues.Add(newValue);
            }

            tracingService.Trace("New and existing multi-select optionset values have been merged correctly. Total options: {0} ", existingValues.Count);
            return existingValues;
        }
    }
}

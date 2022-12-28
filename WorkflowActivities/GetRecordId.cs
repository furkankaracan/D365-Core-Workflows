using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Linq;

namespace D365_Core_Workflows.WorkflowActivities
{
    public class GetRecordId : CodeActivity
    {
        [RequiredArgument]
        [Input("Record URL")]
        [Default("")]
        public InArgument<string> RecordURL { get; set; }

        [Output("Record ID")]
        public OutArgument<string> RecordID { get; set; }

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

            string recordURL = this.RecordURL.Get(executionContext);
            if (recordURL == string.Empty)
                throw new InvalidPluginExecutionException("Invalid input parameters! Please contact with your System Administrator.");

            #endregion Parameters

            string id = GetRecordIdFromURL(recordURL);

            tracingService.Trace($"Setting OutArgument RecordID with {id}");
            this.RecordID.Set(executionContext, id);

        }

        public string GetRecordIdFromURL(string recordURL)
        {
            tracingService.Trace("Started GetRecordIdFromURL");

            string[] urlParts = recordURL.Split("?".ToArray());
            string[] urlParams = urlParts[1].Split("&".ToCharArray());
            string id = urlParams[1].Replace("id=", "");

            tracingService.Trace("Ended GetRecordIdFromURL");
            return id;
        }
    }
}

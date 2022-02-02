using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;

namespace D365_Core_Workflows.WorkflowActivities
{
    public class CloneRecord : CodeActivity
    {
        #region WFA Parameters

        [RequiredArgument]
        [Input("Clonning Record URL")]
        [ReferenceTarget("")]
        public InArgument<String> ClonningRecordURL { get; set; }


        [Input("Prefix")]
        [Default("")]
        public InArgument<String> Prefix { get; set; }

        [Input("Fields to Ignore")]
        [Default("")]
        public InArgument<String> FieldstoIgnore { get; set; }

        [Output("Cloned Guid")]
        public OutArgument<String> ClonedGuid { get; set; }

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

            #region "Read Parameters"
            String clonningRecordURL = this.ClonningRecordURL.Get(executionContext);
            if (clonningRecordURL == null || clonningRecordURL == "")
            {
                throw new InvalidPluginExecutionException("Invalid input parameters! Please contact with your System Administrator.");
            }

            string[] urlParts = clonningRecordURL.Split("?".ToArray());
            string[] urlParams = urlParts[1].Split("&".ToCharArray());

            int objectTypeCode = Convert.ToInt32(urlParams[0].Replace("etc=", ""));

            string entityName = GetEntityNameFromTypeCode(objectTypeCode);
            string objectId = urlParams[1].Replace("id=", "");

            tracingService.Trace($"ObjectTypeCode: {objectTypeCode} -- ParentId:  {objectId}");

            string prefix = this.Prefix.Get(executionContext);
            string fieldstoIgnore = this.FieldstoIgnore.Get(executionContext);
            #endregion

            tracingService.Trace("Starting CloneSelectedRecord");

            var createdGUID = CloneSelectedRecord(entityName, objectId, fieldstoIgnore, prefix);
            ClonedGuid.Set(executionContext, createdGUID.ToString());

            tracingService.Trace("Ending CloneSelectedRecord");

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

        public Guid CloneSelectedRecord(string entityName, string objectId, string fieldstoIgnore, string prefix)
        {
            tracingService.Trace("Starting CloneSelectedRecord");

            if (fieldstoIgnore == null) fieldstoIgnore = "";
            fieldstoIgnore = fieldstoIgnore.ToLower();

            tracingService.Trace("fieldstoIgnore=" + fieldstoIgnore);
            Entity retrievedObject = service.Retrieve(entityName, new Guid(objectId), new ColumnSet(allColumns: true));
            tracingService.Trace("retrieved object OK");

            Entity newEntity = new Entity(entityName);
            string PrimaryIdAttribute = "";
            string PrimaryNameAttribute = "";
            List<string> atts = getEntityAttributesToClone(entityName, service, ref PrimaryIdAttribute, ref PrimaryNameAttribute);

            foreach (string att in atts)
            {
                if (fieldstoIgnore != null && fieldstoIgnore != "")
                {
                    if (Array.IndexOf(fieldstoIgnore.Split(';'), att) >= 0 || Array.IndexOf(fieldstoIgnore.Split(','), att) >= 0)
                    {
                        continue;
                    }
                }

                if (retrievedObject.Attributes.Contains(att) && att != "statuscode" && att != "statecode"
                    || att.StartsWith("partylist-"))
                {
                    if (att.StartsWith("partylist-"))
                    {
                        string att2 = att.Replace("partylist-", "");

                        string fetchParty = @"<fetch version='1.0' output-format='xml - platform' mapping='logical' distinct='true'>
                                                <entity name='activityparty'>
                                                    <attribute name = 'partyid'/>
                                                        <filter type = 'and' >
                                                            <condition attribute = 'activityid' operator= 'eq' value = '" + objectId + @"' />
                                                            <condition attribute = 'participationtypemask' operator= 'eq' value = '" + getParticipation(att2) + @"' />
                                                         </filter>
                                                </entity>
                                            </fetch> ";

                        RetrieveMultipleRequest fetchRequest1 = new RetrieveMultipleRequest
                        {
                            Query = new FetchExpression(fetchParty)
                        };
                        tracingService.Trace(fetchParty);
                        EntityCollection returnCollection = ((RetrieveMultipleResponse)service.Execute(fetchRequest1)).EntityCollection;


                        EntityCollection arrPartiesNew = new EntityCollection();
                        tracingService.Trace("attribute:{0}", att2);

                        foreach (Entity ent in returnCollection.Entities)
                        {
                            Entity party = new Entity("activityparty");
                            EntityReference partyid = (EntityReference)ent.Attributes["partyid"];


                            party.Attributes.Add("partyid", new EntityReference(partyid.LogicalName, partyid.Id));
                            tracingService.Trace("attribute:{0}:{1}:{2}", att2, partyid.LogicalName, partyid.Id.ToString());
                            arrPartiesNew.Entities.Add(party);
                        }

                        newEntity.Attributes.Add(att2, arrPartiesNew);
                        continue;
                    }

                    tracingService.Trace("attribute:{0}", att);
                    if (att == PrimaryNameAttribute && prefix != null)
                    {
                        retrievedObject.Attributes[att] = prefix + retrievedObject.Attributes[att];
                    }
                    newEntity.Attributes.Add(att, retrievedObject.Attributes[att]);
                }
            }

            tracingService.Trace("Starting Creating cloned object...");
            Guid createdGUID = service.Create(newEntity);
            tracingService.Trace("Ending Creating cloned object");

            if (newEntity.Attributes.Contains("statuscode") && newEntity.Attributes.Contains("statecode"))
            {
                Entity record = service.Retrieve(entityName, createdGUID, new ColumnSet("statuscode", "statecode"));

                if (retrievedObject.Attributes["statuscode"] != record.Attributes["statuscode"] ||
                    retrievedObject.Attributes["statecode"] != record.Attributes["statecode"])
                {
                    Entity setStatusEnt = new Entity(entityName, createdGUID);
                    setStatusEnt.Attributes.Add("statuscode", retrievedObject.Attributes["statuscode"]);
                    setStatusEnt.Attributes.Add("statecode", retrievedObject.Attributes["statecode"]);

                    service.Update(setStatusEnt);
                }
            }

            tracingService.Trace("cloned object is finished");
            return createdGUID;
        }

        public List<string> getEntityAttributesToClone(string entityName, IOrganizationService service,
           ref string PrimaryIdAttribute, ref string PrimaryNameAttribute)
        {
            List<string> atts = new List<string>();
            RetrieveEntityRequest req = new RetrieveEntityRequest()
            {
                EntityFilters = EntityFilters.Attributes,
                LogicalName = entityName
            };

            RetrieveEntityResponse res = (RetrieveEntityResponse)service.Execute(req);
            PrimaryIdAttribute = res.EntityMetadata.PrimaryIdAttribute;

            foreach (AttributeMetadata attMetadata in res.EntityMetadata.Attributes)
            {
                if (attMetadata.IsPrimaryName.Value)
                {
                    PrimaryNameAttribute = attMetadata.LogicalName;
                }
                if ((attMetadata.IsValidForCreate.Value || attMetadata.IsValidForUpdate.Value)
                    && !attMetadata.IsPrimaryId.Value)
                {
                    //tracingService.Trace("Tipo:{0}", attMetadata.AttributeTypeName.Value.ToLower());
                    if (attMetadata.AttributeTypeName.Value.ToLower() == "partylisttype")
                    {
                        atts.Add("partylist-" + attMetadata.LogicalName);
                        //atts.Add(attMetadata.LogicalName);
                    }
                    else
                    {
                        atts.Add(attMetadata.LogicalName);
                    }
                }
            }

            return (atts);
        }

        protected string getParticipation(string attributeName)
        {
            string participationType = "";
            switch (attributeName)
            {
                case "from":
                    participationType = "1";
                    break;
                case "to":
                    participationType = "2";
                    break;
                case "cc":
                    participationType = "3";
                    break;
                case "bcc":
                    participationType = "4";
                    break;

                case "organizer":
                    participationType = "7";
                    break;
                case "requiredattendees":
                    participationType = "5";
                    break;
                case "optionalattendees":
                    participationType = "6";
                    break;
                case "customer":
                    participationType = "11";
                    break;
                case "resources":
                    participationType = "10";
                    break;
            }
            return participationType;
        }
    }
}

using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System.Collections.Generic;

namespace CRMAdvanceCustomPlugins
{
    public class DisplayRelatedActivities : IPlugin
    {
        /*
         * Code for DisplayRelatedActivites based on a blog post called "Show ALL related actvities in a subgrid" by Jonas Rapp: https://jonasr.app/2016/04/all-activities/
         * Modified for use in Ellucian CRM Advance by Matt Hatton, University of Wyoming: mhatton@uwyo.edu
        */
        private enum pluginStage
        {
            preValidation = 10,
            preOperation = 20,
            mainOperation = 30,
            postOperation = 40
        }

        public void Execute(IServiceProvider serviceProvider)
        {
            /*
             * Extract the tracing service for us in debugging sandboxed plug-ins.
             * If you are not registering the plug-in in the sandbox, then you do
             * not have to add any tracing service related code.
             */
            ITracingService tracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
#if DEBUG
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = factory.CreateOrganizationService(context.UserId);
#endif
            try
            {
                if (context.MessageName != "RetrieveMultiple" || context.Stage != 20 || context.Mode != 0 ||
                    !context.InputParameters.Contains("Query") || !(context.InputParameters["Query"] is QueryExpression))
                {
                    tracer.Trace("Not expected context");
                    return;
                }
                var query = context.InputParameters["Query"] as QueryExpression;
#if DEBUG
                var fetch1 = ((QueryExpressionToFetchXmlResponse)service.Execute(new QueryExpressionToFetchXmlRequest() { Query = query })).FetchXml;
                tracer.Trace($"Query before:\n{fetch1}");
#endif
                if (ReplaceRegardingCondition(query, tracer, serviceProvider, context))
                {
#if DEBUG
                    var fetch2 = ((QueryExpressionToFetchXmlResponse)service.Execute(new QueryExpressionToFetchXmlRequest() { Query = query })).FetchXml;
                    tracer.Trace($"Query after:\n{fetch2}");
#endif
                    context.InputParameters["Query"] = query;
                }
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException(e.Message);
            }
        }


        private static bool ReplaceRegardingCondition(QueryExpression query, ITracingService tracer, IServiceProvider serviceProvider, IPluginExecutionContext context)
        {
            if (query.EntityName != "activitypointer" || query.Criteria == null || query.Criteria.Conditions == null || query.Criteria.Conditions.Count < 2)
            {
                tracer.Trace("Not expected query");
                return false;
            }

            ConditionExpression nullCondition = null;
            ConditionExpression regardingCondition = null;

            tracer.Trace("Checking criteria for expected conditions");
            foreach (ConditionExpression cond in query.Criteria.Conditions)
            {
                if (cond.AttributeName == "activityid" && cond.Operator == ConditionOperator.Null)
                {
                    tracer.Trace("Found triggering null condition");
                    nullCondition = cond;
                }
                else if (cond.AttributeName == "regardingobjectid" && cond.Operator == ConditionOperator.Equal && cond.Values.Count == 1 && cond.Values[0] is Guid)
                {
                    tracer.Trace("Found condition for regardingobjectid");
                    regardingCondition = cond;
                }
                else
                {
                    tracer.Trace($"Disregarding condition for {cond.AttributeName}");
                }
            }
            if (nullCondition == null || regardingCondition == null)
            {
                tracer.Trace("Missing expected null condition or regardingobjectid condition");
                return false;
            }
            var regardingId = (Guid)regardingCondition.Values[0];
            tracer.Trace($"Found regarding id: {regardingId}");

            
            // Obtain the organization service reference which you will need for web service calls.
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));

            // Perform this CRUD operation as the user who initiated this plugin.
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            // Build the fetch query to search for sposues
            QueryExpression spouseLookup = new QueryExpression("elcn_personalrelationship");
            string[] spouseLookupColumnSet = {"elcn_person1id","elcn_person2id","elcn_personalrelationshipid" };
            spouseLookup.ColumnSet = new ColumnSet(spouseLookupColumnSet);
            FilterExpression spouseFilter = new FilterExpression(LogicalOperator.And);
            spouseFilter.AddCondition("elcn_person1id", ConditionOperator.Equal, regardingId.ToString());
            spouseLookup.Criteria.AddFilter(spouseFilter);

            // Make sure the relationship is spousal
            LinkEntity relationshipTypeLink = new LinkEntity("elcn_personalrelationship", "elcn_personalrelationshiptype", "elcn_relationshiptype1id", "elcn_personalrelationshiptypeid", JoinOperator.Inner);
            FilterExpression relationshipTypeFilter = new FilterExpression(LogicalOperator.And);
            relationshipTypeFilter.AddCondition("elcn_isspousal", ConditionOperator.Equal, System.Boolean.TrueString);
            relationshipTypeLink.LinkCriteria = relationshipTypeFilter;

            // Make sure the spousal relationship is current
            LinkEntity relationshipStatusLink = new LinkEntity("elcn_personalrelationshiptype", "elcn_status", "elcn_relationshipstatusid", "elcn_statusid", JoinOperator.Inner);
            FilterExpression relationshipStatusFilter = new FilterExpression(LogicalOperator.And);
            relationshipStatusFilter.AddCondition("elcn_name", ConditionOperator.Equal, "Current");
            relationshipStatusLink.LinkCriteria = relationshipStatusFilter;
            relationshipTypeLink.LinkEntities.Add(relationshipStatusLink);

            // Add the filters and links to the spouse lookup
            spouseLookup.LinkEntities.Add(relationshipTypeLink);

            // Build the fetch query to search for child organizations
            QueryExpression childOrganizationLookup = new QueryExpression("account");
            string[] childOrganizationLookupColumnSet = { "accountid" };
            childOrganizationLookup.ColumnSet = new ColumnSet(childOrganizationLookupColumnSet);
            childOrganizationLookup.Criteria.AddCondition("accountid", ConditionOperator.Under, regardingId.ToString());

            // All additional ids will go into this list
            List<string> lookupIDs = new List<string>();

            // Add the original regarding id to this list
            lookupIDs.Add(regardingId.ToString());

            try
            {
                // Find and add any current sposue to the lookup list
                EntityCollection spouse = service.RetrieveMultiple(spouseLookup);
                foreach (Entity e in spouse.Entities)
                {
                    Guid spouseGuid = e.GetAttributeValue<EntityReference>("elcn_person2id").Id;
                    string spouseGuidString = spouseGuid.ToString();
                    if (!lookupIDs.Contains(spouseGuidString))
                    {
                        lookupIDs.Add(spouseGuidString);
                    }
                }

                // Find and add any child organizations to the lookup list
                EntityCollection childOrganizations = service.RetrieveMultiple(childOrganizationLookup);
                foreach (Entity o in childOrganizations.Entities)
                {
                    Guid childOrginizationGuid = o.Id;
                    string childOrganizationGuidString = childOrginizationGuid.ToString();
                    if (!lookupIDs.Contains(childOrganizationGuidString))
                    {
                        lookupIDs.Add(childOrganizationGuidString);
                    }
                }
            }
            catch (Exception ex)
            {
                tracer.Trace("Display Related Activities Plugin could not run the related record queries. Error_Msg: {0}", ex.ToString());
                return false;
            }

            // Remove the activity-is-null criteria and the regarding condition
            tracer.Trace("Removing triggering conditions");
            query.Criteria.Conditions.Remove(nullCondition);
            query.Criteria.Conditions.Remove(regardingCondition);

            // Include left outer links to find activities which are communication activities for removal at the end
            LinkEntity letterLink = query.AddLink("letter", "activityid", "activityid", JoinOperator.LeftOuter);
            letterLink.EntityAlias = "removeletter";
            letterLink.LinkCriteria.AddCondition("elcn_communicationactivityid", ConditionOperator.NotNull);

            LinkEntity emailLink = query.AddLink("email", "activityid", "activityid", JoinOperator.LeftOuter);
            emailLink.EntityAlias = "removeemail";
            emailLink.LinkCriteria.AddCondition("elcn_communicationactivityid", ConditionOperator.NotNull);

            LinkEntity phoneLink = query.AddLink("phonecall", "activityid", "activityid", JoinOperator.LeftOuter);
            phoneLink.EntityAlias = "removephone";
            phoneLink.LinkCriteria.AddCondition("elcn_communicationactivityid", ConditionOperator.NotNull);

            FilterExpression removeCommunicationActivities = new FilterExpression(LogicalOperator.And);
            // Remove letter communication activities
            ConditionExpression removeLetterCommunicationActivities = new ConditionExpression("activityid", ConditionOperator.Null);
            removeLetterCommunicationActivities.EntityName = "removeletter";
            // Remove email communication activities
            ConditionExpression removeEmailCommunicationActivities = new ConditionExpression("activityid", ConditionOperator.Null);
            removeEmailCommunicationActivities.EntityName = "removeemail";
            // Remove phone communication activities
            ConditionExpression removePhoneCommunicationActivities = new ConditionExpression("activityid", ConditionOperator.Null);
            removePhoneCommunicationActivities.EntityName = "removephone";

            removeCommunicationActivities.AddCondition(removeLetterCommunicationActivities);
            removeCommunicationActivities.AddCondition(removeEmailCommunicationActivities);
            removeCommunicationActivities.AddCondition(removePhoneCommunicationActivities);
            // Add the new remove communication activities filter
            query.Criteria.AddFilter(removeCommunicationActivities);

            // Add a left outer link to the activity party
            string[] allLookupIds = lookupIDs.ToArray();
            tracer.Trace("Adding link-entity and condition for activity party");
            LinkEntity linkActivityParty = query.AddLink("activityparty", "activityid", "activityid", JoinOperator.LeftOuter);
            linkActivityParty.LinkCriteria.AddCondition("partyid", ConditionOperator.In, allLookupIds);
            // Give the left outer link an alias (entity name) so we can refer to the link in the outer-most filter
            linkActivityParty.EntityAlias = "activity_party";

            // Using the alias of the linkActivityParty look activity party ids or regarding object ids
            FilterExpression partyOrReguarding = new FilterExpression(LogicalOperator.Or);
            ConditionExpression activityParty = new ConditionExpression("activityid", ConditionOperator.NotNull);
            // Reference the alias of the link with EntityName
            activityParty.EntityName = "activity_party";
            partyOrReguarding.AddCondition(activityParty);
            // Add a new regarding filter to include all new ids
            partyOrReguarding.AddCondition("regardingobjectid", ConditionOperator.In, allLookupIds);

            // Add the new party and regarding filter
            query.Criteria.AddFilter(partyOrReguarding);

            // Make sure to get a distinct list because these linked entities will likely bring in duplicate activities
            query.Distinct = true;

            return true;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace GdprRetentionPolicy
{
    public class UpdateAccountOnCaseClosure : IPlugin
    {      
          public const int ProblemSolved = 5;
          public const int InformationProvided = 1000;
          public const int Cancelled = 6;
          public const int Merged = 2000;
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                ITracingService tracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                if (context.MessageName.ToLower() != "update" && context.Stage != 20)
                {
                    return;
                }
                try
                {                   
                    Entity preImage = new Entity();
                    EntityReference accountRef = null;
                    int casesClosed = 0;           

                    Entity targetEntity = context.InputParameters["Target"] as Entity;                  
                    DateTime modifiedOnDate = targetEntity.GetAttributeValue<DateTime>("modifiedon");                 

                    int statusValue = targetEntity.GetAttributeValue<OptionSetValue>("statecode").Value;
                    int statusCodeValue = targetEntity.GetAttributeValue<OptionSetValue>("statuscode").Value;
                    tracer.Trace("Status code is" + statusCodeValue);

                    if (context.PreEntityImages.Contains("PreImage"))
                    {
                        preImage = context.PreEntityImages["PreImage"];
                        accountRef = preImage.GetAttributeValue<EntityReference>("customerid");                       
                    }                

                    Entity accountAttributes = service.Retrieve("account", accountRef.Id, new ColumnSet("tfl_numberofcasesclosed"));

                    if (accountAttributes.Attributes.Contains("tfl_numberofcasesclosed")){
                        casesClosed = accountAttributes.GetAttributeValue<int>("tfl_numberofcasesclosed");
                    }        
                   
                    // Status reason codes for the case resolution & case cancellation
                    if (accountRef != null && (statusCodeValue == ProblemSolved || statusCodeValue == InformationProvided || statusCodeValue == Cancelled || statusCodeValue == Merged))
                    {                        
                        Entity updateAccount = new Entity("account");
                        updateAccount.Id = accountRef.Id;
                        updateAccount["tfl_retentiondate"] = modifiedOnDate.AddYears(3);
                        updateAccount["tfl_numberofcasesclosed"] = casesClosed + 1;
                        service.Update(updateAccount);
                        tracer.Trace("Account's date has been updated with the case resolution:" + statusCodeValue);
                    }                                     
                }
                catch (InvalidPluginExecutionException ex)
                {
                    throw new InvalidPluginExecutionException(ex.ToString());
                }
            }
        }
    }
}

using System;
using Microsoft.Xrm.Sdk;
using System.Text.RegularExpressions;

namespace CRMAdvanceCustomPlugins
{
    public class ScrubHTML : IPlugin
    {
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
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            try
            {
                // The HTML only needs to be removed when the data is being exported to Excel.
                if (context.ParentContext.MessageName != "ExportToExcel")
                {
                    return;
                }

                // The plugin should be registered on the RetrieveMultiple and Retrieve messages in the Post-Operation execution stage
                if (context.MessageName != "RetrieveMultiple" && context.MessageName != "Retrieve")
                {
                    return;
                }

                if (context.Stage != (int)pluginStage.postOperation)
                {
                    return;
                }                
            }
            catch
            {
                return;
            }

            // Changing the output results needs to be done by modifying the OutputParameters object returned by a post-process context.
            if (context.OutputParameters != null)
            {
                try
                {
                    // This is the list of entities returned by the query just before it is passed off to a function which formats the Excel file.
                    EntityCollection output = (EntityCollection)context.OutputParameters["BusinessEntityCollection"];
                    foreach (Entity e in output.Entities)
                    {
                        // This solution is hard-coded for the attributes of Subject and Description from the ActivityPointer (Activities) table.
                        try
                        {
                            String oldSubject = e.GetAttributeValue<string>("subject");
                            if (oldSubject != null && HasHTML(oldSubject) == true)
                            {
                                String newSubject = RemoveHTML(oldSubject);
                                e["subject"] = newSubject;
                            }

                            String oldDescription = e.GetAttributeValue<string>("description");
                            if (oldDescription != null && HasHTML(oldDescription) == true)
                            {
                                String newDescription = RemoveHTML(oldDescription);
                                e["description"] = newDescription;
                            }
                        }
                        catch (Exception ex)
                        {
                            tracingService.Trace("ScrubHTML plugin triggered. Unable to strip HTML tags. Error: {0}", ex.ToString());
                        }
                    }
                    //tracingService.Trace("ScrubHTML plugin triggered. Updating description to {0}", newDescription);
                }
                catch (Exception ex)
                {
                    tracingService.Trace("ScrubHTML plugin triggered. Unable to get OutputParameters. Error: {0}", ex.ToString());
                }
            }
            else
            {
                tracingService.Trace("ScrubHTML plugin triggered. Output Parameters is null.");
            }

            return;
        }


        // Check to see if the old string has one of these common html tags. If it does, return true.
        private bool HasHTML(String oldString)
        {
            bool hasHTML = false;
            if (oldString == null)
            {
                return hasHTML;
            }

            
            if(Regex.IsMatch(oldString, @"<\s*p\s|<\s*body\s|<\s*html\s|<\s*span\s|<\s*div\s|<\s*br\s"))
            {
                hasHTML = true;
            }

            return hasHTML;
        }

        private String RemoveHTML(String oldString)
        {
            if(oldString == null)
            {
                return null;
            }

            // Replace anything between < angle brackets > with nothing. This is a shotgun approach which removes absolutely everything between
            // opened and closed brackets. If you need something with more finesse you can add it here.
            String newString = Regex.Replace(oldString, @"<.*?>", String.Empty, RegexOptions.Multiline | RegexOptions.Singleline);

            // Copied emails often have the HTML-safe versions of brackets and ampersands encoded in the email. 
            // This reverts them back to their original values after the HTML tags have been removed.
            newString = Regex.Replace(newString, "&lt;", "<");
            newString = Regex.Replace(newString, "&gt;", ">");
            newString = Regex.Replace(newString, "&nbsp;", " ");
            newString = Regex.Replace(newString, "&amp;", "&");

            return newString;
        }
    }
}

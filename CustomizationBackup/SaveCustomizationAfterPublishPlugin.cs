using Microsoft.Xrm.Sdk;
using System;

namespace CustomizationBackup
{
    // Message: Publish, PublishAll
    // Primary Entity: none
    // Secondary Entity: none
    // Stage of Execution: Post-Operation
    // Execution Mode: Synchronous
    // Execution Order: 1
    public class SaveCustomizationAfterPublishPlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            string messageName = context.MessageName;

            if (messageName == "Publish" || messageName == "PublishAll")
            {
                CustomizationBackuper backuper = new CustomizationBackuper(serviceProvider);
                backuper.CreateBackup();
            }
            else
            {
                string errorMessage = $@"SaveCustomizationBeforePublishPlugin registered incorrectly.
                                      Ensure that it is registered on post-operation stage and Publish&PublishAll messages";

                throw new InvalidPluginExecutionException(errorMessage);
            }
        }
    }
}
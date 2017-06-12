using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System;

namespace CustomizationBackup
{
    internal class CustomizationBackuper
    {
        private IPluginExecutionContext _context;
        private IOrganizationService _organizatioService;

        public CustomizationBackuper(IServiceProvider serviceProvider)
        {
            _context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            _organizatioService = serviceFactory.CreateOrganizationService(_context.UserId);
        }

        internal void CreateBackup()
        {
            byte[] exportXml = ExportDefaultSolutionZip();
            Guid backupId = CreateCustomizationBackupEntity();
            AttachDefaulSolutionZip(exportXml, backupId);
        }

        private void AttachDefaulSolutionZip(byte[] exportXml, Guid backupId)
        {
            Entity note = new Entity("annotation");
            note["mimetype"] = "application/zip";
            note["filename"] = $"DefaultSolutionBackup_{DateTime.Now.ToString("yyyyMMddHHmm")}.zip";
            note["documentbody"] = Convert.ToBase64String(exportXml);
            note["objectid"] = new EntityReference("aab_customizationbackup", "aab_customizationbackupid", backupId);

            _organizatioService.Create(note);
        }

        private Guid CreateCustomizationBackupEntity()
        {
            Entity backup = new Entity("aab_customizationbackup");
            backup["aab_name"] = $"{_context.MessageName} on {DateTime.Now}";

            if (_context.InputParameters.ContainsKey("ParameterXml"))
            {
                backup["aab_parameterxml"] = _context.InputParameters["ParameterXml"];
            }

            Guid backupId = _organizatioService.Create(backup);
            return backupId;
        }

        private byte[] ExportDefaultSolutionZip()
        {
            ExportSolutionRequest exportRequest = new ExportSolutionRequest();
            exportRequest.ExportAutoNumberingSettings = true;
            exportRequest.ExportCalendarSettings = true;
            exportRequest.ExportCustomizationSettings = true;
            exportRequest.ExportEmailTrackingSettings = true;
            exportRequest.ExportExternalApplications = true;
            exportRequest.ExportGeneralSettings = true;
            exportRequest.ExportIsvConfig = true;
            exportRequest.ExportMarketingSettings = true;
            exportRequest.ExportOutlookSynchronizationSettings = true;
            exportRequest.ExportRelationshipRoles = true;
            exportRequest.ExportSales = true;
            exportRequest.SolutionName = "Default";

            ExportSolutionResponse exportResponse = (ExportSolutionResponse)_organizatioService.Execute(exportRequest);

            byte[] exportXml = exportResponse.ExportSolutionFile;
            return exportXml;
        }
    }
}
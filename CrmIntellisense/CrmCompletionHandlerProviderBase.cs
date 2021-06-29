using CrmIntellisense.Crm;
using CrmIntellisense.Resources;
using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Connection;
using D365DeveloperExtensions.Core.Models;
using D365DeveloperExtensions.Core.UserOptions;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.Xrm.Tooling.Connector;
using System.ComponentModel;

namespace CrmIntellisense
{
    public class CrmCompletionHandlerProviderBase
    {
        public InfoBarModel CreateMetadataInfoBar()
        {
            var text = new InfoBarTextSpan(Resource.Infobar_RetrievingMetadata);
            InfoBarTextSpan[] spans = { text };
            var infoBarModel = new InfoBarModel(spans);

            return infoBarModel;
        }

        public bool IsIntellisenseEnabled(DTE dte)
        {
            var useIntellisense = UserOptionsHelper.GetOption<bool>(UserOptionProperties.UseIntellisense);
            if (!useIntellisense)
                return false;

            bool value = SharedGlobals.GetGlobal<bool>("UseCrmIntellisense", dte, out bool valueFound);
            if (!valueFound)
                return false;

            var isEnabled = value;
            return isEnabled;
        }

        public void GetData(CrmConnect client, InfoBar infoBar)
        {
            if (CrmMetadata.Metadata != null)
            {
                infoBar.HideInfoBar();
                return;
            }

            var bgw = new BackgroundWorker();

            bgw.DoWork += (_, __) => CrmMetadata.GetMetadata(client);

            bgw.RunWorkerCompleted += (_, __) => infoBar.HideInfoBar();

            bgw.RunWorkerAsync();
        }
    }
}
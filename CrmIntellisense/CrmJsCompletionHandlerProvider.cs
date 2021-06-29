using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Connection;
using EnvDTE;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Utilities;
using Microsoft.Xrm.Tooling.Connector;
using System.ComponentModel.Composition;
using System.Collections.Generic;
namespace CrmIntellisense
{
    [Export(typeof(IVsTextViewCreationListener))]
    [Name("CRM JS Token Completion Handler")]
    [ContentType("JavaScript")]
    [ContentType("JScript")]
    //[ContentType("htmlx")]
    [ContentType("TypeScript")]
    [TextViewRole(PredefinedTextViewRoles.Interactive)]
    internal class CrmJsCompletionHandlerProvider : CrmCompletionHandlerProviderBase, IVsTextViewCreationListener
    {
        [Import]
        internal IVsEditorAdaptersFactoryService AdapterService;
        [Import]
        internal ICompletionBroker CompletionBroker { get; set; }
        [Import]
        internal SVsServiceProvider ServiceProvider { get; set; }

        public void VsTextViewCreated(IVsTextView textViewAdapter)
        {
            //This gets executed 1st as each code file is loaded
            if (!(Microsoft.VisualStudio.Shell.ServiceProvider.GlobalProvider.GetService(typeof(DTE)) is DTE dte))
                return;

            CrmConnect client = SharedGlobals.GetGlobal<CrmConnect>("CrmService", dte, out bool found);
            if (client == null || !found)
                return;

            if (!IsIntellisenseEnabled(dte))
                return;

            ITextView textView = AdapterService.GetWpfTextView(textViewAdapter);
            if (textView == null)
                return;

            CrmJsCompletionCommandHandler CreateCommandHandler() => new CrmJsCompletionCommandHandler(textViewAdapter, textView, this);
            textView.Properties.GetOrCreateSingletonProperty(CreateCommandHandler);

            var metadata = SharedGlobals.GetGlobal<List<Completion>>("CrmMetadata", dte, out found);
            if (metadata != null || !found)
                return;

            var infoBar = new InfoBar(false);
            var infoBarModel = CreateMetadataInfoBar();
            infoBar.ShowInfoBar(infoBarModel);

            GetData(client, infoBar);
        }
    }
}
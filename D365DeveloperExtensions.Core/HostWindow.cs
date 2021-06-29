using D365DeveloperExtensions.Core.Models;
using D365DeveloperExtensions.Core.Resources;
using EnvDTE;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Linq;
using System.Text;
using D365DeveloperExtensions.Core.Connection;
namespace D365DeveloperExtensions.Core
{
    public static class HostWindow
    {
        public static string GetCaption(string caption, CrmConnect client)
        {
            var parts = caption.Split('|');

            var sb = new StringBuilder();
            sb.Append(parts[0]);

            var url = Resource.HostWindow_SetCaption_NotConnected;
            try
            {
                url = WebBrowser.GetBaseCrmUrlFomClient(client).Host.ToString();
            }
            catch(Exception ex)
            {
                url = Resource.HostWindow_SetCaption_NotConnected;
            }
               

            if (!string.IsNullOrEmpty(url))
            {
                sb.Append($" | {Resource.HostWindow_SetCaption_ConnectedTo}: ");

                if (url.EndsWith("/"))
                    url = url.TrimEnd('/');

                sb.Append(url);
            }

            var version = client?.Service?.ConnectedOrgVersion.ToString() ?? "";
            if (!string.IsNullOrEmpty(version))
            {
                sb.Append($" | {Resource.HostWindow_SetCaption_Version}: ");
                sb.Append(version);
            }

            return sb.ToString();
        }

        public static bool IsD365DevExWindow(Window window)
        {
            if (window.ObjectKind == null)
                return false;

            var windowGuid = new Guid(StringFormatting.RemoveBracesToUpper(window.ObjectKind));

            return ExtensionConstants.D365DevExToolWindows.Count(w => w.ToolWindowsId == windowGuid) > 0;
        }

        public static ToolWindow GetD365DevExWindow(Window window)
        {
            if (window.ObjectKind == null)
                return null;

            var windowGuid = new Guid(StringFormatting.RemoveBracesToUpper(window.ObjectKind));

            return ExtensionConstants.D365DevExToolWindows.FirstOrDefault(w => w.ToolWindowsId == windowGuid);
        }

        public static bool IsCrmDexWindowOpen(DTE dte)
        {
            foreach (Window window in dte.Windows)
            {
                if (IsD365DevExWindow(window))
                    return true;
            }

            return false;
        }
    }
}
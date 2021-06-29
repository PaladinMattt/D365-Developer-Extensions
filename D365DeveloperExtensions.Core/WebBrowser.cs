﻿using D365DeveloperExtensions.Core.Models;
using D365DeveloperExtensions.Core.UserOptions;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.Xrm.Sdk.Discovery;
using Microsoft.Xrm.Tooling.Connector;
using NLog;
using System;
using System.Linq;

using D365DeveloperExtensions.Core.Connection;
namespace D365DeveloperExtensions.Core
{
    public static class WebBrowser
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static void OpenCrmPage(CrmConnect client, string contentUrl)
        {
            var crmUri = GetBaseCrmUrlFomClient(client);

            var url = new Uri(crmUri, contentUrl);

            OpenPage(url.ToString());
        }

        public static Uri GetBaseCrmUrlFomClient(CrmConnect client)
        {
            var crmUri = client?.Service?.OrganizationWebProxyClient?.Endpoint.Address.Uri;
            return crmUri;
        }

        public static void OpenUrl(string contentUrl)
        {
            OpenPage(contentUrl);
        }

        private static void OpenPage(string contentUrl)
        {
            try
            {
                var useInternalBrowser = UserOptionsHelper.GetOption<bool>(UserOptionProperties.UseInternalBrowser);

                if (!(Package.GetGlobalService(typeof(DTE)) is DTE dte))
                    throw new ArgumentNullException(Resources.Resource.ErrorMessage_ErrorAccessingDTE);

                if (useInternalBrowser) //Internal VS browser
                    dte.ItemOperations.Navigate(contentUrl);
                else                   //User's default browser
                    System.Diagnostics.Process.Start(contentUrl);
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resources.Resource.ErrorMessage_ErrorAccessingDTE, ex);
                throw;
            }
        }
    }
}
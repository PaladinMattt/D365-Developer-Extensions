﻿using D365DeveloperExtensions.Core.Connection;
using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Logging;
using D365DeveloperExtensions.Core.Resources;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.Xrm.Tooling.Connector;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;

namespace D365DeveloperExtensions.Core.Controls
{
    public partial class SolutionsButton
    {
        public SolutionsButton()
        {
            InitializeComponent();
            DataContext = this;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public static readonly DependencyProperty IsConnectedProperty = DependencyProperty.Register(
            "IsConnected", typeof(bool), typeof(SolutionsButton),
            new PropertyMetadata(default(bool), OnIsConnectedChange));

        public bool IsConnected
        {
            get => (bool)GetValue(IsConnectedProperty);

            set
            {
                SetValue(IsConnectedProperty, value);
                OnPropertyChanged();
            }
        }

        private static void OnIsConnectedChange(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var solutionsButton = d as SolutionsButton;
            solutionsButton?.OnIsConnectedChange(e);
        }

        private void OnIsConnectedChange(DependencyPropertyChangedEventArgs e)
        {
            IsConnected = (bool)e.NewValue;
        }

        private void Solutions_OnClick(object sender, RoutedEventArgs e)
        {
            if (!(Package.GetGlobalService(typeof(DTE)) is DTE dte))
                return;
            var client = SharedGlobals.GetGlobal<D365DeveloperExtensions.Core.Connection.CrmConnect>("CrmService", dte, out bool found);
            if (client == null)
            {
                OutputLogger.WriteToOutputWindow(Resource.ErrorMessage_NotConnectedOrg, MessageType.Error);
                return;
            }

            WebBrowser.OpenCrmPage(client, "tools/Solution/home_solution.aspx?etc=7100&sitemappath=Settings|Customizations|nav_solution");
        }

        private void Solutions_OnLoaded(object sender, RoutedEventArgs e)
        {
            if (!(Package.GetGlobalService(typeof(DTE)) is DTE dte))
                return;

            var client = SharedGlobals.GetGlobal<CrmConnect>("CrmService", dte, out bool found);
            if ( client != null)
                IsConnected = client.Service.ConnectedOrgUniqueName != null;
            else
            {
                IsConnected = false;
                SetBinding(IsEnabledProperty, "IsConnected");
            }
        }
    }
}
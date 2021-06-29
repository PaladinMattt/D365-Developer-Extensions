﻿using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Logging;
using D365DeveloperExtensions.Core.Models;
using EnvDTE;
using Microsoft.Xrm.Tooling.Connector;
using NLog;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using WebResourceDeployer.Resources;
using WebResourceDeployer.ViewModels;
using D365DeveloperExtensions.Core.Connection;
namespace WebResourceDeployer
{
    public partial class NewWebResource
    {
        #region Private

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private readonly CrmConnect _client;
        private readonly DTE _dte;
        private ObservableCollection<WebResourceType> _webResourceTypes;
        private readonly Project _selectedProject;

        #endregion

        #region Public

        public Guid NewId;
        public int NewType;
        public string NewName;
        public string NewDisplayName;
        public string NewBoundFile;
        public string NewDescription;
        public Guid NewSolutionId;

        public ObservableCollection<WebResourceType> WebResourceTypes
        {
            get => _webResourceTypes;
            set
            {
                _webResourceTypes = value;
                OnPropertyChanged();
            }
        }

        #endregion

        #region  Events

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion

        public NewWebResource(CrmConnect client, DTE dte, ObservableCollection<ComboBoxItem> projectFiles, Guid selectedSolutionId, Project selectedProject)
        {
            InitializeComponent();
            DataContext = this;
            Owner = Application.Current.MainWindow;

            _client = client;
            _dte = dte;
            _selectedProject = selectedProject;

            var result = GetSolutions(selectedSolutionId);
            if (!result)
            {
                MessageBox.Show(Resource.ErrorMessage_ErrorRetrievingSolutions);
                DialogResult = false;
                Close();
            }

            Files.ItemsSource = projectFiles;
            WebResourceTypes =
                D365DeveloperExtensions.Core.Models.WebResourceTypes.GetTypes(client.Service.ConnectedOrgVersion.Major, false);
        }

        private async void Create_OnClick(object sender, RoutedEventArgs e)
        {
            if (!ValidateForm())
                return;

            var filePath = GetFilePath();
            if (string.IsNullOrEmpty(filePath))
                return;

            var relativePath = ((ComboBoxItem)Files.SelectedItem).Content.ToString();
            var name = UniqueName.Text.Trim();
            var prefix = Prefix.Text;
            var type = ((WebResourceType)Type.SelectedItem).Type;
            var displayName = DisplayName.Text.Trim();
            var description = Description.Text.Trim();

            Overlay.ShowMessage(_dte, $"{Resource.Message_Creating}...");

            var webResource =
                Crm.WebResource.CreateNewWebResourceEntity(type, prefix, name, displayName, description, filePath, _selectedProject);

            var webResourceId = await Task.Run(() => Crm.WebResource.CreateWebResourceInCrm(_client, webResource));
            if (webResourceId == Guid.Empty)
            {
                Overlay.HideMessage(_dte);
                DialogResult = false;
                Close();
            }

            var solution = (CrmSolution)Solutions.SelectedItem;
            if (solution.SolutionId != ExtensionConstants.DefaultSolutionId)
            {
                var addedToSolution = await Task.Run(() => Crm.Solution.AddWebResourceToSolution(_client, solution.UniqueName, webResourceId));
                if (!addedToSolution)
                {
                    Overlay.HideMessage(_dte);
                    DialogResult = false;
                    Close();
                    return;
                }
            }

            NewId = webResourceId;
            NewType = type;
            NewName = prefix + name;
            if (!string.IsNullOrEmpty(displayName))
                NewDisplayName = displayName;
            NewBoundFile = relativePath;
            NewDescription = description;
            NewSolutionId = solution.SolutionId;

            Overlay.HideMessage(_dte);
            DialogResult = true;
            Close();
        }

        private string GetFilePath()
        {
            var projectItem = (ProjectItem)((ComboBoxItem)Files.SelectedItem).Tag;
            var filePath = projectItem.Properties.Item("FullPath").Value.ToString();
            if (File.Exists(filePath))
                return filePath;

            OutputLogger.WriteToOutputWindow($"{Resource.Message_MissingFile}: " + filePath, MessageType.Error);
            MessageBox.Show(Resource.MessageBox_FileDoesNotExist);
            return null;
        }

        private bool ValidateForm()
        {
            if (Crm.WebResource.ValidateName(UniqueName.Text))
                return true;

            MessageBox.Show(Resource.Message_InvalidWebResourceName);
            UniqueName.Focus();
            return false;
        }

        private void Files_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Files.SelectedItem == null)
            {
                UniqueName.Text = string.Empty;
                DisplayName.Text = string.Empty;
                return;
            }

            var fileName = ((ComboBoxItem)Files.SelectedItem).Content.ToString();
            var extensionType = D365DeveloperExtensions.Core.Models.WebResourceTypes.GetExtensionType(fileName);
            if (extensionType == FileExtensionType.Ts)
                fileName = fileName.Replace(".ts", ".js");

            DisplayName.Text = FileNameToDisplayName(fileName);
            var extension = Path.GetExtension(fileName);
            UniqueName.Text = fileName;

            if (string.IsNullOrEmpty(extension))
            {
                Type.SelectedValue = null;
                return;
            }

            switch (extensionType)
            {
                case FileExtensionType.Map:
                    Type.SelectedItem = WebResourceTypes.FirstOrDefault(t => t.Name == FileExtensionType.Xml.ToString().ToUpper());
                    break;
                case FileExtensionType.Ts:
                    Type.SelectedItem = WebResourceTypes.FirstOrDefault(t => t.Name == FileExtensionType.Js.ToString().ToUpper());

                    break;
                default:
                    Type.SelectedItem = WebResourceTypes.FirstOrDefault(t => t.Name == extensionType.ToString().ToUpper());
                    break;
            }
        }

        private string FileNameToDisplayName(string fileName)
        {
            if (fileName.Count(s => s == '/') != 1) //nested in folder
                return fileName;

            fileName = fileName.Replace("/", string.Empty);

            if (fileName.StartsWith(Prefix.Text, StringComparison.InvariantCultureIgnoreCase))
                fileName = fileName.Substring(Prefix.Text.Length, fileName.Length - Prefix.Text.Length);

            return fileName;
        }

        private void Solutions_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Solutions.SelectedItem != null)
            {
                var solution = (CrmSolution)Solutions.SelectedItem;
                Prefix.Text = solution.Prefix + "_";
            }
            else
                Prefix.Text = Resource.DefaultPrefix;
        }

        private bool GetSolutions(Guid selectedSolutionId)
        {
            Overlay.ShowMessage(_dte, $"{Resource.Message_RetrievingSolutions}...");

            var results = Crm.Solution.RetrieveSolutionsFromCrm(_client, false);

            var solutions = ModelBuilder.CreateCrmSolutionView(results);

            if (selectedSolutionId != Guid.Empty)
            {
                var sel = solutions.FindIndex(s => s.SolutionId == selectedSolutionId);
                if (sel != -1)
                    Solutions.SelectedIndex = sel;
            }
            else
                Solutions.SelectedIndex = 0;

            Solutions.ItemsSource = solutions;

            Overlay.HideMessage(_dte);

            return true;
        }

        private void Cancel_OnClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
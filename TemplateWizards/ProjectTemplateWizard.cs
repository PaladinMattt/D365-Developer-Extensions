﻿using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Models;
using D365DeveloperExtensions.Core.UserOptions;
using D365DeveloperExtensions.Core.Vs;
using EnvDTE;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.TemplateWizard;
using NLog;
using NuGet.VisualStudio;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Xml;
using TemplateWizards.Models;
using TemplateWizards.Resources;
using WizardCancelledException = Microsoft.VisualStudio.TemplateWizard.WizardCancelledException;

namespace TemplateWizards
{
    public class ProjectTemplateWizard : IWizard
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private DTE _dte;
        private string _coreVersion;
        private string _clientVersion;
        private string _clientPackage;
        private ProjectType _crmProjectType = ProjectType.Plugin;
        private bool _needsCore;
        private bool _needsWorkflow;
        private bool _needsClient;
        private bool _isUnitTest;
        private bool _signAssembly;
        private string _destDirectory;
        private string _unitTestFrameworkPackage;
        private CustomTemplate _customTemplate;
        private bool _addFile = true;
        private string _typesXrmVersion;

        public void RunStarted(object automationObject, Dictionary<string, string> replacementsDictionary, WizardRunKind runKind, object[] customParams)
        {
            try
            {
                _dte = (DTE)automationObject;

                ProjectDataHandler.AddOrUpdateReplacements("$referenceproject$", "False", ref replacementsDictionary);
                if (replacementsDictionary.ContainsKey("$destinationdirectory$"))
                    _destDirectory = replacementsDictionary["$destinationdirectory$"];

                if (replacementsDictionary.ContainsKey("$wizarddata$"))
                {
                    var wizardData = replacementsDictionary["$wizarddata$"];
                    ReadWizardData(wizardData);
                }

                if (_isUnitTest)
                    PreHandleUnitTestProjects(replacementsDictionary);

                if (_needsCore)
                    PreHandleCrmAssemblyProjects(replacementsDictionary);

                if (_crmProjectType == ProjectType.CustomItem)
                    replacementsDictionary = PreHandleCustomItem(replacementsDictionary);

                if (_crmProjectType == ProjectType.TypeScript)
                    PreHandleTypeScriptProjects();
            }
            catch (WizardBackoutException)
            {
                try
                {
                    var destination = new DirectoryInfo(replacementsDictionary["$destinationdirectory$"]);
                    FileSystem.DeleteDirectory(replacementsDictionary["$destinationdirectory$"]);
                    //Delete solution directory if empty
                    if (destination.Parent != null && FileSystem.IsDirectoryEmpty(replacementsDictionary["$solutiondirectory$"]))
                        FileSystem.DeleteDirectory(replacementsDictionary["$solutiondirectory$"]);
                }
                catch
                {
                    // If it fails (doesn't exist/contains files/read-only), let the directory stay.
                }
                throw;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_TemplateWizardError, ex);
                MessageBox.Show(Resource.ErrorMessage_TemplateWizardError);
                throw new WizardCancelledException(Resource.ErrorMessage_WizardCancelInternalError, ex);
            }
        }

        private Dictionary<string, string> PreHandleCustomItem(Dictionary<string, string> replacementsDictionary)
        {
            var templateFolder = UserOptionsHelper.GetOption<string>(UserOptionProperties.CustomTemplatesPath);
            _addFile = CustomTemplateHandler.ValidateTemplateFolder(templateFolder);
            if (!_addFile)
                return replacementsDictionary;

            _addFile = CustomTemplateHandler.ValidateTemplateFile(templateFolder);
            if (!_addFile)
                return replacementsDictionary;

            var templates = CustomTemplateHandler.GetTemplateConfig(templateFolder);
            if (templates == null)
            {
                _addFile = false;
                return replacementsDictionary;
            }

            var results = CustomTemplateHandler.GetTemplatesByLanguage(templates, "CSharp");
            if (results.Count == 0)
            {
                MessageBox.Show(Resource.MessageBox_AddCustomTemplate);
                _addFile = false;
                return replacementsDictionary;
            }

            var templatePicker = CustomTemplateHandler.GetCustomTemplate(results);
            if (templatePicker.SelectedTemplate == null)
            {
                _addFile = false;
                return replacementsDictionary;
            }

            _customTemplate = templatePicker.SelectedTemplate;

            var content = CustomTemplateHandler.GetTemplateContent(templateFolder, _customTemplate, replacementsDictionary);

            replacementsDictionary.Add("$customtemplate$", content);

            return replacementsDictionary;
        }

        private void PreHandleCrmAssemblyProjects(Dictionary<string, string> replacementsDictionary)
        {
            var sdkVersionPicker = new SdkVersionPicker(_needsWorkflow, _needsClient);
            var result = sdkVersionPicker.ShowModal();
            if (!result.HasValue || result.Value == false)
                throw new WizardBackoutException();

            _coreVersion = sdkVersionPicker.CoreVersion;
            _clientVersion = sdkVersionPicker.ClientVersion;
            _clientPackage = sdkVersionPicker.ClientPackage;

            if (!string.IsNullOrEmpty(_clientVersion))
            {
                ProjectDataHandler.AddOrUpdateReplacements("$useXrmToolingClientUsing$",
                    Versioning.StringToVersion(_clientVersion).Major >= 8 ? "1" : "0", ref replacementsDictionary);
            }

            var coreVersion = Versioning.StringToVersion(_coreVersion);
            var v462BaseVersion = new Version(9, 0, 2, 9);

            if ((_crmProjectType == ProjectType.Console && _clientPackage != Resource.SdkAssemblyExtensions) || coreVersion >= v462BaseVersion)
            {
                var targetFrameworkVersion = Versioning.StringToVersion(replacementsDictionary["$targetframeworkversion$"]);
                if (targetFrameworkVersion < new Version(4, 6, 2))
                    ProjectDataHandler.AddOrUpdateReplacements("$targetframeworkversion$", "4.6.2", ref replacementsDictionary);

                // 4.7.1 is max version for plug-ins & workflows Online
                if (targetFrameworkVersion >= new Version(4, 7, 2) && _crmProjectType == ProjectType.Plugin || _crmProjectType == ProjectType.Workflow)
                    ProjectDataHandler.AddOrUpdateReplacements("$targetframeworkversion$", "4.7.1", ref replacementsDictionary);
            }
            else
                ProjectDataHandler.AddOrUpdateReplacements("$targetframeworkversion$", "4.5.2", ref replacementsDictionary); ;
        }

        private void PreHandleUnitTestProjects(Dictionary<string, string> replacementsDictionary)
        {
            var testProjectPicker = new TestProjectPicker();
            var result = testProjectPicker.ShowModal();
            if (!result.HasValue || result.Value == false)
                throw new WizardBackoutException();

            if (testProjectPicker.SelectedProject != null)
            {
                var solution = _dte.Solution;
                var project = testProjectPicker.SelectedProject;
                var path = string.Empty;
                var projectPath = Path.GetDirectoryName(project.FullName);
                var solutionPath = Path.GetDirectoryName(solution.FullName);
                if (!string.IsNullOrEmpty(projectPath) && !string.IsNullOrEmpty(solutionPath))
                {
                    if (projectPath.StartsWith(solutionPath))
                        path = "..\\" + project.UniqueName;
                    else
                        path = project.FullName;
                }

                ProjectDataHandler.AddOrUpdateReplacements("$referenceproject$", "True", ref replacementsDictionary);
                ProjectDataHandler.AddOrUpdateReplacements("$projectPath$", path, ref replacementsDictionary);
                ProjectDataHandler.AddOrUpdateReplacements("$projectId$", project.Kind, ref replacementsDictionary);
                ProjectDataHandler.AddOrUpdateReplacements("$projectName$", project.Name, ref replacementsDictionary);
            }

            if (testProjectPicker.SelectedUnitTestFramework != null)
            {
                _unitTestFrameworkPackage = testProjectPicker.SelectedUnitTestFramework.NugetName;

                ProjectDataHandler.AddOrUpdateReplacements("$useXrmToolingClientUsing$",
                    testProjectPicker.SelectedUnitTestFramework.CrmMajorVersion >= 8 ? "1" : "0", ref replacementsDictionary);
            }
            else
            {
                if (testProjectPicker.SelectedProject == null)
                    return;

                var version = ProjectWorker.GetSdkCoreVersion(testProjectPicker.SelectedProject);
                ProjectDataHandler.AddOrUpdateReplacements("$useXrmToolingClientUsing$",
                    Versioning.StringToVersion(version).Major >= 8 ? "1" : "0", ref replacementsDictionary);
            }
        }

        private void PreHandleTypeScriptProjects()
        {
            var history = NpmProcessor.GetPackageHistory("@types/xrm");

            if (history == null)
            {
                MessageBox.Show(Resource.MessageBox_NPMError);
                throw new WizardBackoutException();
            }

            var npmPicker = new NpmPicker(history);
            var result = npmPicker.ShowModal();
            if (!result.HasValue || result.Value == false)
                throw new WizardBackoutException();

            _typesXrmVersion = npmPicker.SelectedPackage.Version;
        }

        public bool ShouldAddProjectItem(string filePath)
        {
            return _addFile;
        }

        public void RunFinished()
        {
        }

        public void BeforeOpeningFile(ProjectItem projectItem)
        {
        }

        public void ProjectItemFinishedGenerating(ProjectItem projectItem)
        {
            if (!_addFile)
                return;

            if (_crmProjectType != ProjectType.CustomItem)
                return;

            if (_customTemplate == null)
                return;

            var project = projectItem.ContainingProject;

            CustomTemplateHandler.AddTemplateReferences(_customTemplate, project);

            CustomTemplateHandler.InstallTemplateNuGetPackages(_customTemplate, project);

            if (!string.IsNullOrEmpty(_customTemplate.FileName))
                projectItem.Name = _customTemplate.FileName;
        }

        public void ProjectFinishedGenerating(Project project)
        {
            var componentModel = (IComponentModel)Package.GetGlobalService(typeof(SComponentModel));
            if (componentModel == null)
                return;

            var installer = componentModel.GetService<IVsPackageInstaller>();

            switch (_crmProjectType)
            {
                case ProjectType.UnitTest:
                    PostHandleUnitTestProjects(project, installer);
                    break;
                case ProjectType.Console:
                case ProjectType.Plugin:
                case ProjectType.Workflow:
                    PostHandleCrmAssemblyProjects(project, installer);
                    break;
                case ProjectType.TypeScript:
                    PostHandleTypeScriptProject(project);
                    break;
                case ProjectType.SolutionPackage:
                    PostHandleSolutionPackagerProject(project);
                    break;
            }

            _dte.ExecuteCommand("File.SaveAll");
        }

        private void PostHandleSolutionPackagerProject(Project project)
        {
            foreach (SolutionConfiguration solutionConfiguration in _dte.Solution.SolutionBuild.SolutionConfigurations)
                foreach (SolutionContext solutionContext in solutionConfiguration.SolutionContexts)
                    solutionContext.ShouldBuild = false;

            //Delete bin & obj folders
            FileSystem.DeleteDirectory($"{Path.GetDirectoryName(project.FullName)}//bin");
            FileSystem.DeleteDirectory($"{Path.GetDirectoryName(project.FullName)}//obj");
            project.ProjectItems.AddFolder("package");
        }

        private void PostHandleTypeScriptProject(Project project)
        {
            NpmProcessor.InstallPackage("@types/xrm", _typesXrmVersion, ProjectWorker.GetProjectPath(project), true);

            _dte.ExecuteCommand("ProjectandSolutionContextMenus.CrossProjectMultiItem.RefreshFolder");
        }

        private void PostHandleUnitTestProjects(Project project, IVsPackageInstaller installer)
        {
            NuGetProcessor.InstallPackage(installer, project, ExtensionConstants.MsTestTestAdapter, null);
            NuGetProcessor.InstallPackage(installer, project, ExtensionConstants.MsTestTestFramework, null);

            if (_unitTestFrameworkPackage != null)
                NuGetProcessor.InstallPackage(installer, project, _unitTestFrameworkPackage, null);
        }

        private void PostHandleCrmAssemblyProjects(Project project, IVsPackageInstaller installer)
        {
            try
            {
                project.DTE.SuppressUI = true;

                //Install all the NuGet packages
                project = (Project)((Array)_dte.ActiveSolutionProjects).GetValue(0);
                NuGetProcessor.InstallPackage(installer, project, Resource.SdkAssemblyCore, _coreVersion);
                if (_needsWorkflow)
                    NuGetProcessor.InstallPackage(installer, project, Resource.SdkAssemblyWorkflow, _coreVersion);
                if (_needsClient)
                    NuGetProcessor.InstallPackage(installer, project, _clientPackage, _clientVersion);

                ProjectWorker.ExcludeFolder(project, "bin");
                ProjectWorker.ExcludeFolder(project, "performance");

                if (_signAssembly)
                    Signing.GenerateKey(project, _destDirectory);
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorProcessingTemplate, ex);
                MessageBox.Show(Resource.ErrorMessage_ErrorProcessingTemplate);
            }
        }

        private void ReadWizardData(string wizardData)
        {
            var settings = new XmlReaderSettings { ConformanceLevel = ConformanceLevel.Fragment };

            var el = "";

            using (var reader = XmlReader.Create(new StringReader(wizardData), settings))
                while (reader.Read())
                {
                    switch (reader.NodeType)
                    {
                        case XmlNodeType.Element:
                            el = reader.Name;
                            break;
                        case XmlNodeType.Text:
                            switch (el)
                            {
                                case "CRMProjectType":
                                    _crmProjectType = (ProjectType)Enum.Parse(typeof(ProjectType), reader.Value);
                                    break;
                                case "NeedsCore":
                                    _needsCore = bool.Parse(reader.Value);
                                    break;
                                case "NeedsWorkflow":
                                    _needsWorkflow = bool.Parse(reader.Value);
                                    break;
                                case "NeedsClient":
                                    _needsClient = bool.Parse(reader.Value);
                                    break;
                                case "IsUnitTest":
                                    _isUnitTest = bool.Parse(reader.Value);
                                    break;
                                case "SignAssembly":
                                    _signAssembly = bool.Parse(reader.Value);
                                    break;
                            }
                            break;
                    }
                }
        }
    }
}
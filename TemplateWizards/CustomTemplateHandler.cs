﻿using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Logging;
using EnvDTE;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.TemplateWizard;
using Newtonsoft.Json;
using NLog;
using NuGet.VisualStudio;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using TemplateWizards.Models;
using TemplateWizards.Resources;
using VSLangProj;

namespace TemplateWizards
{
    public class CustomTemplateHandler
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static CustomTemplates GetTemplateConfig(string templateFolder)
        {
            var path = Path.Combine(templateFolder, ExtensionConstants.TemplateConfigFile);
            if (!File.Exists(path))
                return null;

            try
            {
                CustomTemplates templates;
                using (var file = File.OpenText(path))
                {
                    var serializer = new JsonSerializer();
                    templates = (CustomTemplates)serializer.Deserialize(file, typeof(CustomTemplates));
                }

                return templates;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, $"{Resource.ErrorMessage_UnableReadDeserializeConfig}: {path}", ex);
                MessageBox.Show($"{Resource.ErrorMessage_UnableReadDeserializeConfig}: {path}");

                return null;
            }
        }

        public static List<CustomTemplate> GetTemplatesByLanguage(CustomTemplates templates, string language)
        {
            return templates.Templates
                .Where(t => t.Language.Equals(language, StringComparison.InvariantCultureIgnoreCase)).ToList();
        }

        public static string GetTemplateContent(string templateFolder, CustomTemplate template, Dictionary<string, string> replacementsDictionary)
        {
            var path = Path.Combine(templateFolder, template.RelativePath);
            if (!File.Exists(path))
                return null;

            var content = File.ReadAllText(path);

            if (!template.CoreReplacements)
                return content;

            foreach (var keyValuePair in replacementsDictionary)
                content = content.Replace(keyValuePair.Key, keyValuePair.Value);

            return content;
        }

        public static void InstallTemplateNuGetPackages(CustomTemplate customTemplate, Project project)
        {
            if (customTemplate.CustomTemplateNuGetPackages.Count <= 0)
                return;

            var componentModel = (IComponentModel)Package.GetGlobalService(typeof(SComponentModel));
            if (componentModel == null)
                return;

            var installer = componentModel.GetService<IVsPackageInstaller>();

            foreach (var package in customTemplate.CustomTemplateNuGetPackages)
                AddPackage(installer, package, project);
        }

        private static void AddPackage(IVsPackageInstaller installer, CustomTemplateNuGetPackage package, Project project)
        {
            var packageVersion = string.IsNullOrEmpty(package.Version) ?
                null :
                package.Version;

            try
            {
                NuGetProcessor.InstallPackage(installer, project, package.Name, packageVersion);
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, $"{Resource.ErrorMessage_FailedToAddNuGetPackage}: {package.Name} {packageVersion}", ex);
            }
        }

        public static void AddTemplateReferences(CustomTemplate customTemplate, Project project)
        {
            if (customTemplate.CustomTemplateReferences.Count <= 0)
                return;

            if (!(project?.Object is VSProject vsproject))
                return;

            foreach (var reference in customTemplate.CustomTemplateReferences)
                D365DeveloperExtensions.Core.Vs.ProjectWorker.AddProjectReference(vsproject, reference.Name);
        }

        public static string GetTemplateFileTemplate()
        {
            var streamResourceInfo = Application.GetResourceStream(new Uri("TemplateWizards;component/Template/templates.json",
                UriKind.Relative));

            if (streamResourceInfo == null)
                return null;

            var sr = new StreamReader(streamResourceInfo.Stream);
            return sr.ReadToEnd();
        }

        public static void CreateTemplateFileTemplate(string templateFolder)
        {
            var content = GetTemplateFileTemplate();
            var path = Path.Combine(templateFolder, ExtensionConstants.TemplateConfigFile);

            var file = File.Create(path);
            file.Close();

            File.WriteAllText(path, content);
        }

        public static bool ValidateTemplateFolder(string templateFolder)
        {
            if (string.IsNullOrEmpty(templateFolder))
            {
                MessageBox.Show(Resource.MessageBox_MissingTemplateFolder);
                return false;
            }

            if (!Directory.Exists(templateFolder))
            {
                MessageBox.Show(Resource.MessageBox_MissingTemplateFolder);
                return false;
            }

            return true;
        }

        public static bool ValidateTemplateFile(string templateFolder)
        {
            if (File.Exists(Path.Combine(templateFolder, ExtensionConstants.TemplateConfigFile)))
                return true;

            var createResult = MessageBox.Show($"{Resource.MessageBox_CreateConfigFile}: {templateFolder}",
                Resource.MessageBox_Title_ConfirmCreateConfigFile, MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);

            if (createResult != MessageBoxResult.Yes)
                return false;

            CreateTemplateFileTemplate(templateFolder);
            MessageBox.Show(Resource.MessageBox_AddTemplateFiles);

            return false;
        }

        public static CustomTemplatePicker GetCustomTemplate(List<CustomTemplate> results)
        {
            var templatePicker = new CustomTemplatePicker(results);
            var result = templatePicker.ShowModal();
            if (!result.HasValue || result.Value == false)
                throw new WizardBackoutException();

            if (templatePicker.SelectedTemplate == null)
                OutputLogger.WriteToOutputWindow(Resource.ErrorMessage_NoSelectedTemplate, MessageType.Error);

            return templatePicker;
        }
    }
}
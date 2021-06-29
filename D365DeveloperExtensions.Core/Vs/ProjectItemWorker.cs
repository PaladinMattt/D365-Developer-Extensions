﻿using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;

namespace D365DeveloperExtensions.Core.Vs
{
    public static class ProjectItemWorker
    {
        private static readonly IEnumerable<string> FileKinds = new[] { VSConstants.GUID_ItemType_PhysicalFile.ToString() };
        private static readonly IEnumerable<string> FolderKinds = new[] { VSConstants.GUID_ItemType_PhysicalFolder.ToString() };
        private static readonly char[] PathSeparatorChars = { Path.DirectorySeparatorChar };
        private static readonly Dictionary<string, string> KnownNestedFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
            { "web.debug.config", "web.config" },
            { "web.release.config", "web.config" }
        };

        public static void ProcessProjectItem(IVsSolution solutionService, Project project)
        {
            //https://www.mztools.com/articles/2014/MZ2014006.aspx
            if (solutionService.GetProjectOfUniqueName(project.UniqueName, out var projectHierarchy) != VSConstants.S_OK)
                return;

            if (projectHierarchy == null)
                return;

            foreach (ProjectItem projectItem in project.ProjectItems)
            {
                string fileFullName = null;

                try
                {
                    fileFullName = projectItem.FileNames[0];
                }
                catch
                {
                    // ignored
                }

                if (string.IsNullOrEmpty(fileFullName))
                    continue;

                if (projectHierarchy.ParseCanonicalName(fileFullName, out var itemId) == VSConstants.S_OK)
                    MessageBox.Show($"File: {fileFullName}\r\nItem Id: 0x{itemId:X}");
            }
        }

        public static uint GetProjectItemId(IVsSolution solutionService, string projectUniqueName, ProjectItem projectItem)
        {
            if (solutionService.GetProjectOfUniqueName(projectUniqueName, out var projectHierarchy) != VSConstants.S_OK)
                return uint.MinValue;

            if (projectHierarchy == null)
                return uint.MinValue;

            string fileFullName;

            try
            {
                fileFullName = projectItem.FileNames[0];
            }
            catch
            {
                return uint.MinValue;
            }

            if (string.IsNullOrEmpty(fileFullName))
                return uint.MinValue;

            return projectHierarchy.ParseCanonicalName(fileFullName, out var itemId) == VSConstants.S_OK
                ? itemId
                : uint.MinValue;
        }

        public static ProjectItem GetProjectItemFromItemId(IVsSolution solutionService, string projectUniqueName, uint projectItemId)
        {
            if (solutionService.GetProjectOfUniqueName(projectUniqueName, out var projectHierarchy) != VSConstants.S_OK)
                return null;

            if (projectHierarchy == null)
                return null;

            projectHierarchy.GetProperty(projectItemId, (int)__VSHPROPID.VSHPROPID_ExtObject, out var objProjectItem);

            var projectItem = objProjectItem as ProjectItem;

            return projectItem;
        }

        public static string CreateValidFolderName(string name)
        {
            string[] illegal = { "/", "?", ":", "&", "\\", "*", "\"", "<", ">", "|", "#", "_" };
            var rxString = string.Join("|", illegal.Select(Regex.Escape));
            name = Regex.Replace(name, rxString, string.Empty);

            return name;
        }

        public static ProjectItem GetProjectItem(Project project, string path)
        {
            var folderPath = Path.GetDirectoryName(path);
            var itemName = Path.GetFileName(path);

            var container = GetProjectItems(project, folderPath);

            if (container == null || !container.TryGetFile(itemName, out var projectItem) && !container.TryGetFolder(itemName, out projectItem))
                return null;

            return projectItem;
        }

        private static ProjectItem GetProjectItem(ProjectItems projectItems, string name, IEnumerable<string> allowedItemKinds)
        {
            try
            {
                var projectItem = projectItems.Item(name);
                if (projectItem != null && allowedItemKinds.Contains(projectItem.Kind, StringComparer.OrdinalIgnoreCase))
                    return projectItem;
            }
            catch
            {
                //ignored
            }

            return null;
        }

        public static ProjectItems GetProjectItems(Project project, string folderPath, bool createIfNotExists = false)
        {
            if (string.IsNullOrEmpty(folderPath))
                return project.ProjectItems;

            var pathParts = folderPath.Split(PathSeparatorChars, StringSplitOptions.RemoveEmptyEntries);

            object cursor = project;

            var fullPath = project.GetFullPath();
            var folderRelativePath = string.Empty;

            foreach (var part in pathParts)
            {
                fullPath = Path.Combine(fullPath, part);
                folderRelativePath = Path.Combine(folderRelativePath, part);

                cursor = GetOrCreateFolder(cursor, fullPath, part, createIfNotExists);
                if (cursor == null)
                    return null;
            }

            return GetProjectItems(cursor);
        }

        private static ProjectItems GetProjectItems(object parent)
        {
            if (parent is Project project)
                return project.ProjectItems;

            if (parent is ProjectItem projectItem)
                return projectItem.ProjectItems;

            return null;
        }

        public static bool TryGetFolder(this ProjectItems projectItems, string name, out ProjectItem projectItem)
        {
            projectItem = GetProjectItem(projectItems, name, FolderKinds);

            return projectItem != null;
        }

        public static bool TryGetFile(this ProjectItems projectItems, string name, out ProjectItem projectItem)
        {
            projectItem = GetProjectItem(projectItems, name, FileKinds);

            if (projectItem == null)
                return TryGetNestedFile(projectItems, name, out projectItem);

            return projectItem != null;
        }

        private static bool TryGetNestedFile(ProjectItems projectItems, string name, out ProjectItem projectItem)
        {
            if (!KnownNestedFiles.TryGetValue(name, out var parentFileName))
                parentFileName = Path.GetFileNameWithoutExtension(name);

            var parentProjectItem = GetProjectItem(projectItems, parentFileName, FileKinds);

            projectItem = parentProjectItem != null ?
                GetProjectItem(parentProjectItem.ProjectItems, name, FileKinds) :
                null;

            return projectItem != null;
        }

        public static string GetFullPath(this Project project)
        {
            var fullPath = project.GetPropertyValue<string>("FullPath");
            if (string.IsNullOrEmpty(fullPath))
                return fullPath;

            if (File.Exists(fullPath))
                fullPath = Path.GetDirectoryName(fullPath);

            return fullPath;
        }

        public static T GetPropertyValue<T>(this Project project, string propertyName)
        {
            if (project.Properties == null)
                return default(T);

            try
            {
                var property = project.Properties.Item(propertyName);
                if (property != null)
                    return (T)property.Value;
            }
            catch (ArgumentException)
            {
                //ignored
            }

            return default(T);
        }

        // 'parentItem' can be either a Project or ProjectItem
        private static ProjectItem GetOrCreateFolder(object parentItem, string fullPath, string folderName, bool createIfNotExists)
        {
            if (parentItem == null)
                return null;

            var projectItems = GetProjectItems(parentItem);
            if (projectItems.TryGetFolder(folderName, out var subFolder))
                return subFolder;

            if (!createIfNotExists)
                return null;

            try
            {
                return projectItems.AddFromDirectory(fullPath);
            }
            catch (NotImplementedException)
            {
                // This is the case for F#'s project system, we can't add from directory so we fall back to this impl
                return projectItems.AddFolder(folderName);
            }
        }

        public static string GetRelativePath(ProjectItem projectItem)
        {
            var project = projectItem.ContainingProject;
            var fullName = projectItem.Properties.Item("FullPath").Value.ToString();
            var relativePath = fullName.Replace(ProjectWorker.GetProjectPath(project), string.Empty);
            if (relativePath.StartsWith("\\"))
                relativePath = relativePath.TrimStart('\\');

            return relativePath;
        }

        public static string GetRelativePathFromPath(string projectPath, string filePath)
        {
            var relativePath = filePath.Replace(projectPath, string.Empty);
            if (relativePath.StartsWith("\\"))
                relativePath = relativePath.TrimStart('\\');

            return relativePath;
        }
    }
}
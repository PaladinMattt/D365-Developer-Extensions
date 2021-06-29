﻿using D365DeveloperExtensions.Core.Resources;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;

namespace D365DeveloperExtensions.Core.Vs
{
    public static class SolutionWorker
    {
        public static void SetBuildConfigurationOff(SolutionConfigurations buildConfigurations, string projectName)
        {
            foreach (SolutionConfiguration buildConfiguration in buildConfigurations)
            {
                if (!string.Equals(buildConfiguration.Name, Resource.Constant_ProjectBuildConfig_Debug, StringComparison.CurrentCultureIgnoreCase) &&
                    !string.Equals(buildConfiguration.Name, Resource.Constant_ProjectBuildConfig_Release, StringComparison.CurrentCultureIgnoreCase))
                    continue;

                var contexts = buildConfiguration.SolutionContexts;
                foreach (SolutionContext solutionContext in contexts)
                {
                    if (solutionContext.ProjectName == projectName)
                        solutionContext.ShouldBuild = false;
                }
            }
        }

        public static IList<Project> GetProjects()
        {
            if (!(Package.GetGlobalService(typeof(DTE)) is DTE dte))
                return null;

            var projects = dte.Solution.Projects;
            var list = new List<Project>();
            var item = projects.GetEnumerator();
            while (item.MoveNext())
            {
                var project = (Project)item.Current;
                if (project == null)
                    continue;

                switch (project.Kind)
                {
                    case ProjectKinds.vsProjectKindSolutionFolder:
                        list.AddRange(GetSolutionFolderProjects(project));
                        break;
                    case Constants.vsProjectKindMisc:
                        continue;
                    default:
                        list.Add(project);
                        break;
                }
            }

            return list;
        }

        private static IEnumerable<Project> GetSolutionFolderProjects(Project solutionFolder)
        {
            var list = new List<Project>();
            for (var i = 1; i <= solutionFolder.ProjectItems.Count; i++)
            {
                var subProject = solutionFolder.ProjectItems.Item(i).SubProject;
                if (subProject == null)
                    continue;

                // If this is another solution folder, do a recursive call, otherwise add
                if (subProject.Kind == ProjectKinds.vsProjectKindSolutionFolder)
                    list.AddRange(GetSolutionFolderProjects(subProject));
                else
                    list.Add(subProject);
            }
            return list;
        }
    }
}
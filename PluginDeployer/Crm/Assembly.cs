﻿using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Connection;
using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Logging;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using NLog;
using PluginDeployer.Resources;
using PluginDeployer.Spkl;
using PluginDeployer.ViewModels;
using System;
using ExLogger = D365DeveloperExtensions.Core.Logging.ExtensionLogger;
using Logger = NLog.Logger;

namespace PluginDeployer.Crm
{
    public class Assembly
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static Entity RetrieveAssemblyFromCrm(CrmConnect client, string assemblyName)
        {
            try
            {
                var query = new QueryExpression
                {
                    EntityName = "pluginassembly",
                    ColumnSet = new ColumnSet("pluginassemblyid", "name", "version"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "name",
                                Operator = ConditionOperator.Equal,
                                Values = { assemblyName }
                            }
                        }
                    }
                };

                var assemblies = client.Service.RetrieveMultiple(query);

                if (assemblies.Entities.Count <= 0)
                    return null;

                ExLogger.LogToFile(Logger, $"{Resource.Message_RetrievedAssembly}: {assemblies.Entities[0].Id}", LogLevel.Info);
                OutputLogger.WriteToOutputWindow($"{Resource.Message_RetrievedAssembly}: {assemblies.Entities[0].Id}", MessageType.Info);
                return assemblies.Entities[0];

            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorRetrievingAssembly, ex);

                return null;
            }
        }

        public static Guid UpdateCrmAssembly(CrmConnect client, CrmAssembly crmAssembly)
        {
            try
            {
                var assembly = new Entity("pluginassembly")
                {
                    ["content"] = Convert.ToBase64String(FileSystem.GetFileBytes(crmAssembly.AssemblyPath)),
                    ["name"] = crmAssembly.Name,
                    ["culture"] = crmAssembly.Culture,
                    ["version"] = crmAssembly.Version,
                    ["publickeytoken"] = crmAssembly.PublicKeyToken,
                    ["sourcetype"] = new OptionSetValue(0), // database
                    ["isolationmode"] = crmAssembly.IsolationMode == IsolationModeEnum.Sandbox
                    ? new OptionSetValue(2) // 2 = sandbox
                    : new OptionSetValue(1) // 1= none
                };

                if (crmAssembly.AssemblyId == Guid.Empty)
                {
                    var newId = client.Service.Create(assembly);
                    ExLogger.LogToFile(Logger, $"{Resource.Message_CreatedAssembly}: {crmAssembly.Name}|{newId}", LogLevel.Info);
                    OutputLogger.WriteToOutputWindow($"{Resource.Message_CreatedAssembly}: {crmAssembly.Name}|{newId}", MessageType.Info);
                    return newId;
                }

                assembly.Id = crmAssembly.AssemblyId;
                client.Service.Update(assembly);

                ExLogger.LogToFile(Logger, $"{Resource.Message_UpdatedAssembly}: {crmAssembly.Name}|{crmAssembly.AssemblyId}", LogLevel.Info);
                OutputLogger.WriteToOutputWindow($"{Resource.Message_UpdatedAssembly}: {crmAssembly.Name}|{crmAssembly.AssemblyId}", MessageType.Info);

                return crmAssembly.AssemblyId;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorCreatingOrUpdatingAssembly, ex);

                return Guid.Empty;
            }
        }

        public static bool AddAssemblyToSolution(CrmConnect client, Guid assemblyId, string uniqueName)
        {
            try
            {
                var scRequest = new AddSolutionComponentRequest
                {
                    ComponentType = 91,
                    SolutionUniqueName = uniqueName,
                    ComponentId = assemblyId
                };

                client.Service.Execute(scRequest);

                ExLogger.LogToFile(Logger, $"{Resource.Message_AssemblyAddedSolution}: {uniqueName} - {assemblyId}", LogLevel.Info);
                OutputLogger.WriteToOutputWindow($"{Resource.Message_AssemblyAddedSolution}: {uniqueName} - {assemblyId}", MessageType.Info);

                return true;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorAddingAssemblySolution, ex);

                return false;
            }
        }

        public static bool IsAssemblyInSolution(CrmConnect client, string assemblyName, string uniqueName)
        {
            try
            {
                var query = new FetchExpression($@"<fetch>
                                                          <entity name='solutioncomponent'>
                                                            <attribute name='solutioncomponentid'/>
                                                            <link-entity name='pluginassembly' from='pluginassemblyid' to='objectid'>
                                                              <attribute name='pluginassemblyid'/>
                                                              <filter type='and'>
                                                                <condition attribute='name' operator='eq' value='{assemblyName}'/>
                                                              </filter>
                                                            </link-entity>
                                                            <link-entity name='solution' from='solutionid' to='solutionid'>
                                                              <attribute name='solutionid'/>
                                                              <filter type='and'>
                                                                <condition attribute='uniquename' operator='eq' value='{uniqueName}'/>
                                                              </filter>
                                                            </link-entity>
                                                          </entity>
                                                        </fetch>");

                var results = client.Service.RetrieveMultiple(query);

                var inSolution = results.Entities.Count > 0;

                ExLogger.LogToFile(Logger, $"{Resource.Message_AssemblyInSolution}: {uniqueName} - {assemblyName} - {inSolution}", LogLevel.Info);
                OutputLogger.WriteToOutputWindow($"{Resource.Message_AssemblyInSolution}: {uniqueName} - {assemblyName} - {inSolution}", MessageType.Info);

                return inSolution;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorCheckingAssemblyInSolution, ex);

                return true;
            }
        }
    }
}
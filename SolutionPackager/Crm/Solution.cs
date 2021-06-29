﻿using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Connection;
using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Logging;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using NLog;
using SolutionPackager.Resources;
using SolutionPackager.ViewModels;
using System;
using System.Threading.Tasks;
using ExLogger = D365DeveloperExtensions.Core.Logging.ExtensionLogger;
using Logger = NLog.Logger;
using Task = System.Threading.Tasks.Task;

namespace SolutionPackager.Crm
{
    public static class Solution
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static EntityCollection RetrieveSolutionsFromCrm(CrmConnect client)
        {
            try
            {
                var query = new QueryExpression
                {
                    EntityName = "solution",
                    ColumnSet = new ColumnSet("friendlyname", "solutionid", "uniquename", "version"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "isvisible",
                                Operator = ConditionOperator.Equal,
                                Values = {true}
                            },
                            new ConditionExpression
                            {
                                AttributeName = "ismanaged",
                                Operator = ConditionOperator.Equal,
                                Values = { false }
                            }
                        }
                    },
                    LinkEntities =
                    {
                        new LinkEntity
                        {
                            LinkFromEntityName = "solution",
                            LinkFromAttributeName = "publisherid",
                            LinkToEntityName = "publisher",
                            LinkToAttributeName = "publisherid",
                            Columns = new ColumnSet("customizationprefix"),
                            EntityAlias = "publisher"
                        }
                    },
                    Orders =
                    {
                        new OrderExpression
                        {
                            AttributeName = "friendlyname",
                            OrderType = OrderType.Ascending
                        }
                    }
                };

                var solutions = client.Service.RetrieveMultiple(query);

                ExLogger.LogToFile(Logger, Resource.Message_RetrievedSolutions, LogLevel.Info);
                OutputLogger.WriteToOutputWindow(Resource.Message_RetrievedSolutions, MessageType.Info);

                return solutions;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorRetrievingSolutions, ex);

                return null;
            }
        }

        public static async Task<string> GetSolutionFromCrm(CrmConnect client, CrmSolution selectedSolution, bool managed)
        {
            try
            {
                // Hardcode connection timeout to one-hour to support large solutions.
                //TODO: Find a better way to handle this
                if (client.Service.OrganizationServiceProxy != null)
                    client.Service.OrganizationServiceProxy.Timeout = new TimeSpan(1, 0, 0);
                if (client.Service.OrganizationWebProxyClient != null)
                    client.Service.OrganizationWebProxyClient.InnerChannel.OperationTimeout = new TimeSpan(1, 0, 0);

                var request = new ExportSolutionRequest
                {
                    Managed = managed,
                    SolutionName = selectedSolution.UniqueName
                };
                var response = await Task.Run(() => (ExportSolutionResponse)client.Service.Execute(request));

                ExLogger.LogToFile(Logger, Resource.Message_RetrievedSolution, LogLevel.Info);
                OutputLogger.WriteToOutputWindow(Resource.Message_RetrievedSolution, MessageType.Info);

                var fileName = FileHandler.FormatSolutionVersionString(selectedSolution.UniqueName, selectedSolution.Version, managed);
                var tempFile = FileHandler.WriteTempFile(fileName, response.ExportSolutionFile);

                return tempFile;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorRetrievingSolution, ex);

                return null;
            }
        }

        public static bool ImportSolution(CrmConnect client, string path)
        {
            var solutionBytes = FileSystem.GetFileBytes(path);
            if (solutionBytes == null)
                return false;

            try
            {
                var request = new ImportSolutionRequest
                {
                    CustomizationFile = solutionBytes,
                    OverwriteUnmanagedCustomizations = true,
                    PublishWorkflows = true,
                    ImportJobId = Guid.NewGuid()
                };

                client.Service.Execute(request);

                ExLogger.LogToFile(Logger, $"{Resource.Message_ImportedSolution}: {path}", LogLevel.Info);
                OutputLogger.WriteToOutputWindow($"{Resource.Message_ImportedSolution}: {path}", MessageType.Info);

                return true;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorImportingSolution, ex);

                return false;
            }
        }
    }
}
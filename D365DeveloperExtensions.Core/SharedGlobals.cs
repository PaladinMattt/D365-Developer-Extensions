using EnvDTE;
using Microsoft.VisualStudio.Shell;
using NLog;
using System;

namespace D365DeveloperExtensions.Core
{
    public static class SharedGlobals
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        //CrmService (CrmConnect) = Active connection to CRM
        //UseCrmIntellisense (boolean) = Is CRM intellisense on/off

        public static T GetGlobal<T>(string globalName, DTE dte, out bool valueFound)
        {
            valueFound = false;
            try
            {
                if (dte == null)
                {
                    if (!(Package.GetGlobalService(typeof(DTE)) is DTE dte2))
                        throw new ArgumentNullException(Resources.Resource.ErrorMessage_ErrorAccessingDTE);
                    dte = dte2;
                }

                var globals = dte.Globals;
                valueFound = globals.VariableExists[globalName];
                object val = valueFound  ? globals[globalName] : null;
                if (val is T)
                {
                    return (T)val;
                }

                bool isNullable = (Nullable.GetUnderlyingType(typeof(T)) != null);
                if (!isNullable)
                {
                    return (T)Convert.ChangeType(val, typeof(T));
                }
                else
                {
                    return (T)Convert.ChangeType(val, Nullable.GetUnderlyingType(typeof(T)));
                }
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resources.Resource.ErrorMessage_ErrorAccessingDTE, ex);
                return default(T);
            }
        }

        public static void SetGlobal<T>(string globalName, T value, DTE dte = null)
        {
            try
            {
                if (dte == null)
                {
                    if (!(Package.GetGlobalService(typeof(DTE)) is DTE dte2))
                        throw new ArgumentNullException(Resources.Resource.ErrorMessage_ErrorAccessingDTE);
                    dte = dte2;
                };

                var globals = dte.Globals;
                globals[globalName] = value;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resources.Resource.ErrorMessage_ErrorAccessingDTE, ex);
                throw;
            }
        }
    }
}
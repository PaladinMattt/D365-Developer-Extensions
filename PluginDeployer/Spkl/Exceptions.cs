using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PluginDeployer.Spkl
{
    [Serializable]
    public class SparkleTaskException : Exception
    {
        public enum ExceptionTypes {
           DUPLICATE_STEP,
            DUPLICATE_FILE,
            CONFIG_NOTFOUND,
            NO_TASK_SUPPLIED,
            AUTH_ERROR,
            UTILSNOTFOUND,
            MISSING_PREFIX,
            IMPORT_ERROR,
            CRMSVCUTIL_ERROR,
            SOLUTIONPACKAGER_ERROR,
            NO_WEBRESOURCES_FOUND
        }

        public ExceptionTypes ExceptionType { get; protected set; }
        public SparkleTaskException(ExceptionTypes exectionType, string message) :base (message)
        {
            ExceptionType = exectionType;
        }
    }
}

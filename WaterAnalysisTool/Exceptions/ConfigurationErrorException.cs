using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WaterAnalysisTool.Exceptions
{
    /* 
     * Desc: ConfigurationErrorException occurs when a configuration file cannot be read correctly.
     * Example: Exception would occur if a section of a config file is missing
     */

    [Serializable]
    class ConfigurationErrorException : Exception
    {
        public ConfigurationErrorException() { }

        public ConfigurationErrorException(String msg) : base(msg) { }

        public ConfigurationErrorException(String msg, Exception inner) : base(msg, inner) { }
    }
}

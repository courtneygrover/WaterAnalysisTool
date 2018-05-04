using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WaterAnalysisTool.Exceptions
{
    /*
     * Desc: Exception occurs in a loader when a parser does not hand off parsed data to the loader correctly.
     * Example: Parse() is called and it hands off an empty list of elements to the loader.
     */

    [Serializable]
    class ParseErrorException : Exception
    {
        public ParseErrorException() { }

        public ParseErrorException(String msg) : base(msg) { }

        public ParseErrorException(String msg, Exception inner) : base(msg, inner) { }
    }
}

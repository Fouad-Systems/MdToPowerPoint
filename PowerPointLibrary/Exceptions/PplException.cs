using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.Exceptions
{
    public class PplException : Exception
    {
        public PplException(string message) : base(message)
        {
        }
    }
}

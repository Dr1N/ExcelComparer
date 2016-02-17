using System;

namespace ExcelComparer
{
    class ComparerException : Exception
    {
        public ComparerException(string msg) : base(msg) { } 
    }
}

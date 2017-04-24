using ezResx.Data;
using System;
using System.Collections.Generic;

namespace ezResx.Errors
{
    internal class DataLossException : Exception
    {
        public IEnumerable<ResourceItem> MissingData { get; private set; }
        public DataLossException(IEnumerable<ResourceItem> xlsxRes)
        {
            MissingData = new List<ResourceItem>(xlsxRes);
        }

        public DataLossException(string message, IEnumerable<ResourceItem> xlsxRes) : base(message)
        {
            MissingData = new List<ResourceItem>(xlsxRes);
        }

        public DataLossException(string message, IEnumerable<ResourceItem> xlsxRes, Exception innerException) : base(message, innerException)
        {
            MissingData = new List<ResourceItem>(xlsxRes);
        } 
    }

    internal class InvalidDataException : Exception
    {
        public InvalidDataException()
        {
        }

        public InvalidDataException(string message) : base(message)
        {
        }

        public InvalidDataException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
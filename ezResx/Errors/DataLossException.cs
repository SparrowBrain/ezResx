using ezResx.Data;
using System;
using System.Collections.Generic;

namespace ezResx.Errors
{
    internal class DataLossException : Exception
    {
        public List<ResourceItem> MissingData { get; set; }
        public DataLossException(IEnumerable<ResourceItem> xlsxRes)
        {
            MissingData = new List<ResourceItem>();
            foreach (var lostResource in xlsxRes)
            {
                MissingData.Add(lostResource);
            }

        }

        public DataLossException(string message, IEnumerable<ResourceItem> xlsxRes) : base(message)
        {
            MissingData = new List<ResourceItem>();
            foreach (var lostResource in xlsxRes)
            {
                MissingData.Add(lostResource);
            }
        }

        public DataLossException(string message, IEnumerable<ResourceItem> xlsxRes, Exception innerException) : base(message, innerException)
        {
            MissingData = new List<ResourceItem>();
            foreach (var lostResource in xlsxRes)
            {
                MissingData.Add(lostResource);
            }
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
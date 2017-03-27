using ezResx.Data;
using System;
using System.Collections.Generic;

namespace ezResx.Errors
{
    internal class DataLossException : Exception
    {
        public List<ResourceItem> missingData { get; set; }
        public DataLossException(List<ResourceItem> xlsxRes)
        {
            missingData = new List<ResourceItem>();
            foreach (var lostResource in xlsxRes)
            {
                missingData.Add(lostResource);
            }

        }

        public DataLossException(string message, List<ResourceItem> xlsxRes) : base(message)
        {
            missingData = new List<ResourceItem>();
            foreach (var lostResource in xlsxRes)
            {
                missingData.Add(lostResource);
            }
        }

        public DataLossException(string message, List<ResourceItem> xlsxRes, Exception innerException) : base(message, innerException)
        {
            missingData = new List<ResourceItem>();
            foreach (var lostResource in xlsxRes)
            {
                missingData.Add(lostResource);
            }
        } 

        

        //public void PrintMissingData(List<ResourceItem> missingData)
        //{
        //    foreach (var lostResource in missingData)
        //    {
        //        Console.WriteLine($"{ lostResource.Key.File} {lostResource.Key.Name }");
        //    }
        //}
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
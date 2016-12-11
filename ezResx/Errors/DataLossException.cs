using System;

namespace ezResx.Errors
{
    internal class DataLossException : Exception
    {
        public DataLossException()
        {
        }

        public DataLossException(string message) : base(message)
        {
        }

        public DataLossException(string message, Exception innerException) : base(message, innerException)
        {
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
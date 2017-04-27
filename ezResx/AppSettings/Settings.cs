using System;
using System.Configuration;

namespace ezResx.AppSettings
{
    class Settings
    {
        public string[] ReadAppSettings()
        {
            var settings = ConfigurationManager.AppSettings["PriorityLanguages"];
            string[] appSettings = settings.Split(',');
           return appSettings; 
        }
    }
}

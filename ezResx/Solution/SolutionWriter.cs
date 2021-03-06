﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Resources;
using System.Xml.Linq;
using ezResx.Data;
using ezResx.Errors;

namespace ezResx.Solution
{
    internal class SolutionWriter : SolutionService
    {
        public void AddResourcesToSolution(string solutionPath, List<ResourceItem> resources)
        {
            var projects = GetSolutionProjects(solutionPath);
            var resourcesByProject = resources.GroupBy(x => x.Key.Project);
            foreach (var projectGroup in resourcesByProject)
            {
                var projectPath = projects.FirstOrDefault(x => x.RelativeToSolution.Equals(projectGroup.Key, StringComparison.InvariantCultureIgnoreCase));
                if (projectPath == null)
                {
                    throw new Exception($"Project {projectGroup.Key} deos not exist in solution");
                }

                var project = new Project(projectPath.FullPath);

                var projectDirectory = project.DirectoryPath;


                var resourcesByFile = projectGroup.ToList().GroupBy(x => x.Key.File);
                foreach (var fileGroup in resourcesByFile)
                {
                    var lostData = new List<ResourceItem>();
                    var defaultFilePath = Path.Combine(projectDirectory, fileGroup.Key);
                    if (!File.Exists(defaultFilePath))
                    {
                        throw new Exception($"File {defaultFilePath} does not exist");
                    }


                    // Edit xml
                    //

                    var defaultFile = XElement.Load(defaultFilePath);
                    var defaultElements = defaultFile.Elements("data").ToDictionary(x => x.Attribute("name")?.Value);

                    foreach (var resource in fileGroup.ToList())
                    {
                        var element =
                            defaultFile.Elements("data").FirstOrDefault(x => x.Attribute("name")?.Value == resource.Key.Name);
                        if (element == null)
                        {
                            lostData.Add(resource);
                            continue;
                        }

                        var valueElement = element.Element("value");
                        if (valueElement == null)
                        {
                            throw new Exception(
                                $"No value element exists for {resource.Key.Name} in {defaultFilePath}");
                        }

                        valueElement.Value = resource.Values["default-culture"];
                    }

                    defaultFile.Save(defaultFilePath);

                    // locales

                    var locales = fileGroup.ToList().SelectMany(x => x.Values.Keys.Where(y => y != "default-culture")).Distinct();
                    foreach (var locale in locales)
                    {
                        var localeFilePath = fileGroup.Key.Insert(fileGroup.Key.LastIndexOf('.'), "." + locale);
                        var filePath = Path.Combine(projectDirectory, localeFilePath);
                        
                        if (!File.Exists(filePath))
                        {
                            using (var resxWriter = new ResXResourceWriter(filePath))
                            {
                                foreach (var resource in fileGroup.ToList())
                                {
                                    string value;
                                    if (!resource.Values.TryGetValue(locale, out value))
                                    {
                                        continue;
                                    }

                                    if (!defaultElements.ContainsKey(resource.Key.Name))
                                    {
                                        throw new Exception($"Name {resource.Key.Name} does not exist in {defaultFilePath}, but exists in {filePath}");
                                    }

                                    resxWriter.AddResource(resource.Key.Name, resource.Values[locale]);
                                }
                                resxWriter.Generate();
                            }

                            InlcudeMissingFileIntoProject(project, localeFilePath);

                            continue;
                        }

                        var resourceFile = XElement.Load(filePath);

                        foreach (var resource in fileGroup.ToList())
                        {
                            if (!defaultElements.ContainsKey(resource.Key.Name))
                            {
                                if (!lostData.Contains((resource)))
                                {
                                    lostData.Add(resource);
                                }
                            }

                            var element =
                                resourceFile.Elements("data")
                                    .FirstOrDefault(x => x.Attribute("name")?.Value == resource.Key.Name);


                            string value;
                            if (resource.Values.TryGetValue(locale, out value))
                            {
                                // update
                                if (element == null)
                                {
                                    var newElement = new XElement("data");
                                    newElement.SetAttributeValue("name", resource.Key.Name);
                                    newElement.SetAttributeValue(XNamespace.Xml + "space", "preserve");
                                    var valueElement = new XElement("value");
                                    valueElement.SetValue(resource.Values[locale]);
                                    newElement.Add(valueElement);
                                    resourceFile.Add(newElement);
                                }
                                else
                                {
                                    var valueElement = element.Element("value");
                                    if (valueElement == null)
                                    {
                                        throw new Exception(
                                            $"No value element exists for {resource.Key.Name} in {filePath}");
                                    }

                                    valueElement.Value = resource.Values[locale];
                                }
                            }
                            else
                            {
                                //delete
                                if (element == null)
                                {
                                    continue;
                                }

                                Console.WriteLine($"Removing resource {resource.Key.Name} from {filePath}");
                                element.Remove();
                            }
                        }

                        if (lostData.Any())
                        {
                            throw new DataLossException($"Name does not exist in {defaultFilePath}", lostData);
                        }

                        resourceFile.Save(filePath);

                        InlcudeMissingFileIntoProject(project, localeFilePath);
                    }
                }
            }
        }

        private void InlcudeMissingFileIntoProject(Project project, string localeFilePath)
        {
            if (project.Items.All(x => x.Include != localeFilePath))
            {
                var metadata = new Dictionary<string, string> {{"SubType", "Designer"}};
                project.AddItem("EmbeddedResource", localeFilePath, metadata);
                project.Save();
            }
        }
    }
}
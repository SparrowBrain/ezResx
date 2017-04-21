using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ezResx.Data;

namespace ezResx.Solution
{
    internal class SolutionReader : SolutionService
    {
        public List<ResourceItem> GetSolutionResources(string solutionPath)
        {
            var projects = GetSolutionProjects(solutionPath);

            var resourceList = new List<ResourceItem>();
            foreach (var projectPath in projects)
            {
                var project = new Project(projectPath.FullPath);

                var items = project.Items.Where(x => x.Include.EndsWith(".resx"));

                foreach (var item in items)
                {
                    var filePath = Path.Combine(project.DirectoryPath, item.Include);
                    var withoutExtension = Path.GetFileNameWithoutExtension(item.Include);
                    var locale = withoutExtension.Contains('.') ? withoutExtension.Split('.').Last() : "default-culture";
                    if (locale != "default-culture")
                    {
                        continue;
                    }

                    var file = item.Include;

                    foreach (var dataElement in XElement.Load(filePath).Elements("data"))
                    {
                        if (dataElement.Attribute("type")?.Value == "System.Resources.ResXFileRef, System.Windows.Forms")
                        {
                            continue;
                        }

                        var nameAttribute = dataElement.Attribute("name");
                        if (nameAttribute == null)
                        {
                            throw new Exception($"Data element does not contain name attribute in {filePath}");
                        }

                        var resourceKey = new ResourceKey
                        {
                            Project = projectPath.RelativeToSolution,
                            File = file,
                            Name = nameAttribute.Value
                        };

                        var resourceItem = new ResourceItem
                        {
                            Key = resourceKey,
                            Values = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase)
                        };

                        var valueElement = dataElement.Element("value");
                        if (valueElement == null)
                        {
                            throw new Exception(
                                $"Data element {nameAttribute.Value} does not have a value element in {filePath}");
                        }

                        resourceItem.Values[locale] = valueElement.Value;

                        resourceList.Add(resourceItem);
                    }
                }


                foreach (var item in items)
                {
                    var filePath = Path.Combine(project.DirectoryPath, item.Include);
                    var withoutExtension = Path.GetFileNameWithoutExtension(item.Include);
                    var locale = withoutExtension.Contains('.') ? withoutExtension.Split('.').Last() : "default-culture";

                    if (locale == "default-culture")
                    {
                        continue;
                    }

                    var match = Regex.Match(item.Include, @"([^\.]+)\.[^\.]+\.resx");
                    var file = match.Groups[1].Value + ".resx";

                    foreach (var dataElement in XElement.Load(filePath).Elements("data"))
                    {
                        if (dataElement.Attribute("type")?.Value == "System.Resources.ResXFileRef, System.Windows.Forms")
                        {
                            continue;
                        }

                        var nameAttribute = dataElement.Attribute("name");
                        if (nameAttribute == null)
                        {
                            throw new Exception($"Data element does not contain name attribute in {filePath}");
                        }

                        var resourceKey = new ResourceKey
                        {
                            Project = projectPath.RelativeToSolution,
                            File = file,
                            Name = nameAttribute.Value
                        };

                        var resourceItem = resourceList.FirstOrDefault(x => Equals(x.Key, resourceKey));
                        if (resourceItem == null)
                        {
                            throw new Exception($"No invariant culture resource found for {resourceKey.File} {resourceKey.Name}");
                        }

                        var valueElement = dataElement.Element("value");
                        if (valueElement == null)
                        {
                            throw new Exception(
                                $"Data element {nameAttribute.Value} does not have a value element in {filePath}");
                        }

                        resourceItem.Values[locale] = valueElement.Value;
                    }
                }
            }
            return resourceList;
        }
    }
}
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace ezResx.Solution
{
    internal class Project
    {
        private readonly XNamespace _ns = "http://schemas.microsoft.com/developer/msbuild/2003";
        private readonly XDocument _document;
        private readonly string _fileName;
        public IEnumerable<ProjectItem> Items { get; private set; }
        public string DirectoryPath { get; private set; }

        public Project(string fileName)
        {
            _fileName = fileName;
            DirectoryPath = Path.GetDirectoryName(_fileName);
            var resxItems = new List<ProjectItem>();

            _document = XDocument.Load(_fileName);
            if (_document.Root == null)
            {
                throw new Exception("Invalid project file");
            }

            foreach (var itemGroup in _document.Root.Elements(_ns + "ItemGroup"))
            {
                var items =
                    itemGroup.Elements()
                        .Where(x => x.Attribute("Include") != null)
                        .Select(x => new ProjectItem(x.Name.LocalName, x.Attribute("Include").Value));

                resxItems.AddRange(items);
            }

            Items = resxItems;
        }

        public void AddItem(string itemType, string unevaluatedInclude, IEnumerable<KeyValuePair<string, string>> metadata)
        {
            var item = new ProjectItem(itemType, unevaluatedInclude, metadata);

            var items = new List<ProjectItem>(Items) {item};
            Items = items;

            // XML change
            
            var groupToModify = _document.Root.Elements(_ns + "ItemGroup").FirstOrDefault(itemGroup => itemGroup.Elements(_ns + itemType).Any());

            if (groupToModify == null)
            {
                groupToModify = new XElement(_ns + "ItemGroup");
                _document.Root.Elements(_ns + "ItemGroup").First().AddAfterSelf(groupToModify);
            }

            // Add to group
            var itemToAdd = new XElement(_ns + itemType);
            itemToAdd.SetAttributeValue("Include", unevaluatedInclude);
            foreach (var metaItem in metadata)
            {
                itemToAdd.Add(new XElement(_ns + metaItem.Key, metaItem.Value));
            }

            groupToModify.Add(itemToAdd);
        }

        public void Save()
        {
            _document.Save(_fileName);
        }
    }
}
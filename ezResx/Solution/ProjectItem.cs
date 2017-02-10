using System.Collections.Generic;

namespace ezResx.Solution
{
    internal class ProjectItem
    {
        public string ItemType { get; }

        public string Include { get; }

        public IEnumerable<KeyValuePair<string, string>> Metadata { get; } = new List<KeyValuePair<string, string>>();

        public ProjectItem(string itemType, string include)
        {
            ItemType = itemType;
            Include = include;
        }

        public ProjectItem(string itemType, string include, IEnumerable<KeyValuePair<string, string>> metadata) : this(itemType, include)
        {
            Metadata = metadata;
        }
    }
}
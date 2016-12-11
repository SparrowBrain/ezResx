using System.Collections.Generic;

namespace ezResx.Data
{
    internal class ResourceItem
    {
        public ResourceKey Key { get; set; }
        
        public IDictionary<string, string> Values { get; set; }
    }
}
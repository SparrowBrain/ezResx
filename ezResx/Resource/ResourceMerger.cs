using System.Collections.Generic;
using System.Linq;
using ezResx.Data;
using ezResx.Errors;
using System;

namespace ezResx.Resource
{
    internal class ResourceMerger
    {
        public List<ResourceItem> MergeResources(List<ResourceItem> solutionResources, List<ResourceItem> xlsxResources)
        {
            var xlsxRes = new List<ResourceItem>(xlsxResources);
            foreach (var solutionResource in solutionResources)
            {
                var xlsxResource = xlsxRes.FirstOrDefault(x => x.Key.Equals(solutionResource.Key));
                if (xlsxResource == null)
                {
                    continue;
                }

                foreach (var value in xlsxResource.Values)
                {
                    solutionResource.Values[value.Key] = value.Value;
                }

                xlsxRes.Remove(xlsxResource);
            }

            if (xlsxRes.Count == 0)
            {
                return solutionResources;
            }

            throw new DataLossException("Translations will be lost.", xlsxRes);
        }
    }
}
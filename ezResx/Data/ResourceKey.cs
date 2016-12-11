namespace ezResx.Data
{
    internal class ResourceKey
    {
        public string Project { get; set; }

        public string File { get; set; }

        public string Name { get; set; }

        public override bool Equals(object obj)
        {
            var other = obj as ResourceKey;
            return other != null && Equals(other);
        }

        protected bool Equals(ResourceKey other)
        {
            return string.Equals(Project, other.Project) && string.Equals(File, other.File) && string.Equals(Name, other.Name);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = Project?.GetHashCode() ?? 0;
                hashCode = (hashCode*397) ^ (File?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (Name?.GetHashCode() ?? 0);
                return hashCode;
            }
        }
    }
}
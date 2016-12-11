using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using ezResx.Data;

namespace ezResx.Solution
{
    internal abstract class SolutionService
    {
        protected List<ProjectPath> GetSolutionProjects(string solutionPath)
        {
            var solutionDirectory = Path.GetDirectoryName(solutionPath);
            var solutionRegex = new Regex(@"\u0022([^\u0022]+\.csproj)\u0022");
            string solution;
            using (var reader = new StreamReader(solutionPath))
            {
                solution = reader.ReadToEnd();
            }

            var projects = new List<ProjectPath>();
            var matches = solutionRegex.Matches(solution);
            foreach (Match match in matches)
            {
                var projectPath = match.Groups[1].Value;
                projects.Add(new ProjectPath
                {
                    FullPath = Path.Combine(solutionDirectory, projectPath),
                    RelativeToSolution = projectPath
                });
            }
            return projects;
        }
    }
}
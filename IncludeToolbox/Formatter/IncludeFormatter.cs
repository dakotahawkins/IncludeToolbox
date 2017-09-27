using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace IncludeToolbox.Formatter
{
    public static class IncludeFormatter
    {
        public static string FormatPath(string absoluteIncludeFilename,
                                        FormatterOptionsPage.PathMode pathFormat,
                                        FormatterOptionsPage.UseFileRelativePathMode useFileRelativePathMode,
                                        IEnumerable<string> includeDirectories,
                                        string documentDir,
                                        string includeRootDirectory = null)
        {
            if (pathFormat == FormatterOptionsPage.PathMode.Absolute)
            {
                return absoluteIncludeFilename;
            }

            string fileRelativeIncludeFilename = null;
            if (useFileRelativePathMode == FormatterOptionsPage.UseFileRelativePathMode.Always)
            {
                fileRelativeIncludeFilename = Utils.MakeRelative(documentDir, absoluteIncludeFilename);
            }
            else if ((useFileRelativePathMode == FormatterOptionsPage.UseFileRelativePathMode.OnlyInSameDirectory &&
                      Path.Combine(documentDir, Path.GetFileName(absoluteIncludeFilename)) == absoluteIncludeFilename) ||
                     (useFileRelativePathMode == FormatterOptionsPage.UseFileRelativePathMode.OnlyInSameOrSubDirectory &&
                      absoluteIncludeFilename.StartsWith(documentDir)))
            {
                fileRelativeIncludeFilename = absoluteIncludeFilename.Remove(0, documentDir.Length);
            }

            if (pathFormat == FormatterOptionsPage.PathMode.ForceRelativeToParentDirWithFile &&
                null != absoluteIncludeFilename &&
                null != includeRootDirectory)
            {
                return Utils.MakeRelative(includeRootDirectory, absoluteIncludeFilename);
            }

            if (pathFormat == FormatterOptionsPage.PathMode.Shortest ||
                pathFormat == FormatterOptionsPage.PathMode.Shortest_AvoidUpSteps)
            {
                // todo: Treat std library files special?
                if (absoluteIncludeFilename != null)
                {
                    string bestCandidate = fileRelativeIncludeFilename;

                    foreach (string includeDirectory in includeDirectories)
                    {
                        string proposal = Utils.MakeRelative(includeDirectory, absoluteIncludeFilename);

                        if (proposal.Length < (bestCandidate?.Length ?? Int32.MaxValue))
                        {
                            if (pathFormat == FormatterOptionsPage.PathMode.Shortest ||
                                (proposal.IndexOf("../") < 0 && proposal.IndexOf("..\\") < 0))
                            {
                                bestCandidate = proposal;
                            }
                        }
                    }

                    return bestCandidate;
                }
            }

            return null;
        }

        /// <summary>
        /// Formats the paths of a given list of include line info.
        /// </summary>
        private static void FormatPaths(IEnumerable<IncludeLineInfo> lines,
                                        FormatterOptionsPage.PathMode pathFormat,
                                        FormatterOptionsPage.UseFileRelativePathMode useFileRelativePathMode,
                                        IEnumerable<string> includeDirectories,
                                        string documentDir,
                                        string includeRootDirectory = null)
        {
            if (pathFormat == FormatterOptionsPage.PathMode.Unchanged)
                return;

            foreach (var line in lines)
            {
                string absoluteIncludePath = null;
                bool resolvedPathAsFileRelative = false;
                {
                    string fileRelativeAbsoluteIncludePath =
                        line.TryResolveInclude(new string[] { documentDir }, out resolvedPathAsFileRelative);
                    if (resolvedPathAsFileRelative)
                    {
                        absoluteIncludePath = fileRelativeAbsoluteIncludePath;
                    }
                    else
                    {
                        string includeDirAbsoluteIncludePath =
                            line.TryResolveInclude(includeDirectories, out bool resolvedPath);
                        if (resolvedPath)
                        {
                            absoluteIncludePath = includeDirAbsoluteIncludePath;
                        }
                    }
                }

                if (null != absoluteIncludePath)
                {
                    if (pathFormat == FormatterOptionsPage.PathMode.ForceRelativeToParentDirWithFile &&
                        includeRootDirectory != null &&
                        !absoluteIncludePath.StartsWith(includeRootDirectory))
                    {
                        // Skip files that are outside of the root directory tree (e.g. system includes)
                        continue;
                    }

                    // TODO: from code review
                    // Not very happy with this. The combination of
                    // FormatterOptionsPage.IgnoreFileRelativeMode.InSameDirectory/InSameDirectoryOrSubdirectory
                    // with FormatterOptionsPage.PathMode.Shortest would no longer find the shortest possible path!
                    // We need to pass on the FormatterOptionsPage.IgnoreFileRelativeMode to the
                    // inner FormatPath to be able to find the best path according to the PathMode
                    //
                    // The main take away of my comments is that I think it's best to think of
                    // PathMode as a policy (out of the given paths which one is the best) and of
                    // AllowRelativePathMode / IgnoreFileRelativeMode as a constraint (which paths
                    // are we not allowed to choose).
                    var currentPathFormat = pathFormat;
                    var currentIncludeRootDirectory = includeRootDirectory;
                    string includeFilename = Path.GetFileName(absoluteIncludePath);
                    if ((useFileRelativePathMode == FormatterOptionsPage.UseFileRelativePathMode.OnlyInSameDirectory &&
                         Path.Combine(documentDir, includeFilename) == absoluteIncludePath) ||
                        (useFileRelativePathMode == FormatterOptionsPage.UseFileRelativePathMode.OnlyInSameOrSubDirectory &&
                         absoluteIncludePath.StartsWith(documentDir)))
                    {
                        currentPathFormat = FormatterOptionsPage.PathMode.ForceRelativeToParentDirWithFile;
                        currentIncludeRootDirectory = documentDir;
                    }

                    line.IncludeContent = FormatPath(absoluteIncludePath,
                                                     pathFormat,
                                                     useFileRelativePathMode,
                                                     includeDirectories,
                                                     documentDir,
                                                     currentIncludeRootDirectory) ?? line.IncludeContent;
                }
            }
        }

        private static void FormatDelimiters(IEnumerable<IncludeLineInfo> lines, FormatterOptionsPage.DelimiterMode delimiterMode)
        {
            switch (delimiterMode)
            {
                case FormatterOptionsPage.DelimiterMode.AngleBrackets:
                    foreach (var line in lines)
                        line.SetDelimiterType(IncludeLineInfo.DelimiterType.AngleBrackets);
                    break;
                case FormatterOptionsPage.DelimiterMode.Quotes:
                    foreach (var line in lines)
                        line.SetDelimiterType(IncludeLineInfo.DelimiterType.Quotes);
                    break;
            }
        }

        private static void FormatSlashes(IEnumerable<IncludeLineInfo> lines, FormatterOptionsPage.SlashMode slashMode)
        {
            switch (slashMode)
            {
                case FormatterOptionsPage.SlashMode.ForwardSlash:
                    foreach (var line in lines)
                        line.IncludeContent = line.IncludeContent.Replace('\\', '/');
                    break;
                case FormatterOptionsPage.SlashMode.BackSlash:
                    foreach (var line in lines)
                        line.IncludeContent = line.IncludeContent.Replace('/', '\\');
                    break;
            }
        }

        private static List<IncludeLineInfo> SortIncludes(IList<IncludeLineInfo> lines, FormatterOptionsPage settings, string documentName)
        {
            string[] precedenceRegexes = RegexUtils.FixupRegexes(settings.PrecedenceRegexes, documentName);

            List<IncludeLineInfo> outSortedList = new List<IncludeLineInfo>(lines.Count);

            IEnumerable<IncludeLineInfo> includeBatch;
            int numConsumedItems = 0;

            do
            {
                // Fill in all non-include items between batches.
                var nonIncludeItems = lines.Skip(numConsumedItems).TakeWhile(x => !x.ContainsActiveInclude);
                numConsumedItems += nonIncludeItems.Count();
                outSortedList.AddRange(nonIncludeItems);

                // Process until we hit a preprocessor directive that is not an include.
                // Those are boundaries for the sorting which we do not want to cross.
                includeBatch = lines.Skip(numConsumedItems).TakeWhile(x => x.ContainsActiveInclude || !x.ContainsPreProcessorDirective);
                numConsumedItems += includeBatch.Count();

            } while (SortIncludeBatch(settings, precedenceRegexes, outSortedList, includeBatch) && numConsumedItems != lines.Count);

            return outSortedList;
        }

        private static bool SortIncludeBatch(FormatterOptionsPage settings, string[] precedenceRegexes,
                                            List<IncludeLineInfo> outSortedList, IEnumerable<IncludeLineInfo> includeBatch)
        {
            // Get enumerator and cancel if batch is empty.
            if (!includeBatch.Any())
                return false;

            // Fetch settings.
            FormatterOptionsPage.TypeSorting typeSorting = settings.SortByType;
            bool regexIncludeDelimiter = settings.RegexIncludeDelimiter;
            bool blankAfterRegexGroupMatch = settings.BlankAfterRegexGroupMatch;

            // Select only valid include lines and sort them. They'll stay in this relative sorted
            // order when rearranged by regex precedence groups.
            var includeLines = includeBatch
                .Where(x => x.ContainsActiveInclude)
                .OrderBy(x => x.IncludeContent)
                .ToList();

            if (settings.RemoveDuplicates)
            {
                HashSet<string> uniqueIncludes = new HashSet<string>();
                includeLines.RemoveAll(x => !x.ShouldBePreserved &&
                                            !uniqueIncludes.Add(x.GetIncludeContentWithDelimiters()));
            }

            // Group the includes by the index of the precedence regex they match, or
            // precedenceRegexes.Length for no match, and sort the groups by index.
            var includeGroups = includeLines
                .GroupBy(x =>
                {
                    var includeContent = regexIncludeDelimiter ? x.GetIncludeContentWithDelimiters() : x.IncludeContent;
                    for (int precedence = 0; precedence < precedenceRegexes.Count(); ++precedence)
                    {
                        if (Regex.Match(includeContent, precedenceRegexes[precedence]).Success)
                            return precedence;
                    }

                    return precedenceRegexes.Length;
                }, x => x)
                .OrderBy(x => x.Key);

            // Optional newlines between regex match groups
            var groupStarts = new HashSet<IncludeLineInfo>();
            if (blankAfterRegexGroupMatch && precedenceRegexes.Length > 0 && includeLines.Count() > 1)
            {
                // Set flag to prepend a newline to each group's first include
                foreach (var grouping in includeGroups)
                    groupStarts.Add(grouping.First());
            }

            // Flatten the groups
            var sortedIncludes = includeGroups.SelectMany(x => x.Select(y => y));

            // Sort by angle or quoted delimiters if either of those options were selected
            if (typeSorting == FormatterOptionsPage.TypeSorting.AngleBracketsFirst)
                sortedIncludes = sortedIncludes.OrderBy(x => x.LineDelimiterType == IncludeLineInfo.DelimiterType.AngleBrackets ? 0 : 1);
            else if (typeSorting == FormatterOptionsPage.TypeSorting.QuotedFirst)
                sortedIncludes = sortedIncludes.OrderBy(x => x.LineDelimiterType == IncludeLineInfo.DelimiterType.Quotes ? 0 : 1);

            // Merge sorted includes with original non-include lines
            var sortedIncludeEnumerator = sortedIncludes.GetEnumerator();
            var sortedLines = includeBatch.Select(originalLine =>
            {
                if (originalLine.ContainsActiveInclude)
                {
                    // Replace original include with sorted includes
                    return sortedIncludeEnumerator.MoveNext() ? sortedIncludeEnumerator.Current : new IncludeLineInfo();
                }
                return originalLine;
            });

            if (settings.RemoveEmptyLines)
            {
                // Removing duplicates may have introduced new empty lines
                sortedLines = sortedLines.Where(sortedLine => !string.IsNullOrWhiteSpace(sortedLine.RawLine));
            }

            // Finally, update the actual lines
            {
                bool firstLine = true;
                foreach (var sortedLine in sortedLines)
                {
                    // Handle prepending a newline if requested, as long as:
                    // - this include is the begin of a new group
                    // - it's not the first line
                    // - the previous line isn't already a non-include
                    if (groupStarts.Contains(sortedLine) && !firstLine && outSortedList[outSortedList.Count - 1].ContainsActiveInclude)
                    {
                        outSortedList.Add(new IncludeLineInfo());
                    }
                    outSortedList.Add(sortedLine);
                    firstLine = false;
                }
            }

            return true;
        }

        /// <summary>
        /// Formats all includes in a given piece of text.
        /// </summary>
        /// <param name="text">Text to be parsed for includes.</param>
        /// <param name="documentName">Path to the document the edit is occuring in.</param>
        /// <param name="includeDirectories">A list of include directories</param>
        /// <param name="settings">Settings that determine how the formating should be done.</param>
        /// <returns>Formated text.</returns>
        public static string FormatIncludes(string text, string documentPath, IEnumerable<string> includeDirectories, FormatterOptionsPage settings)
        {
            string documentDir = Utils.GetExactPathName(Path.GetDirectoryName(documentPath));
            string documentName = Path.GetFileNameWithoutExtension(documentPath);
            string includeRootDirectory = null;
            if (settings.PathFormat == FormatterOptionsPage.PathMode.ForceRelativeToParentDirWithFile &&
                !String.IsNullOrWhiteSpace(settings.FromParentDirWithFile))
            {
                var dir = new DirectoryInfo(documentDir);
                while (dir != null)
                {
                    if (File.Exists(Path.Combine(dir.FullName, settings.FromParentDirWithFile)))
                    {
                        includeRootDirectory = Utils.GetExactPathName(dir.FullName);
                        break;
                    }

                    try
                    {
                        dir = dir.Parent;
                    }
                    catch (System.Security.SecurityException)
                    {
                        // Permission denied
                        break;
                    }
                }
            }

            string newLineChars = Utils.GetDominantNewLineSeparator(text);

            var lines = IncludeLineInfo.ParseIncludes(text, settings.RemoveEmptyLines ? ParseOptions.RemoveEmptyLines : ParseOptions.None);

            // Format.
            FormatPaths(lines,
                        settings.PathFormat,
                        settings.UseFileRelativePath,
                        includeDirectories,
                        documentDir,
                        includeRootDirectory);
            FormatDelimiters(lines, settings.DelimiterFormatting);
            FormatSlashes(lines, settings.SlashFormatting);

            // Sorting. Ignores non-include lines.
            lines = SortIncludes(lines, settings, documentName);

            // Combine again.
            return string.Join(newLineChars, lines.Select(x => x.RawLine));
        }
    }
}

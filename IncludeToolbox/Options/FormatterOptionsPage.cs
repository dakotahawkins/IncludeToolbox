using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;

namespace IncludeToolbox
{
    [Guid("B822F53B-32C0-4560-9A84-2F9DA7AB0E4C")]
    public class FormatterOptionsPage : OptionsPage
    {
        public const string SubCategory = "Include Formatter";
        public const string CollectionName = "IncludeFormatter";

        #region Path

        public enum PathMode
        {
            Unchanged,
            Shortest,
            Shortest_AvoidUpSteps,
            Absolute,
            ForceRelativeToParentDirWithFile
        }
        [Category("Path")]
        [DisplayName("Mode")]
        [Description("Changes the path mode to the given pattern.")]
        public PathMode PathFormat { get; set; } = PathMode.ForceRelativeToParentDirWithFile;

        [Category("Path")]
        [DisplayName("Filename for ForceRelativeToParentDirWithFile mode")]
        [Description("The ForceRelativeToParentDirWithFile mode will look for this file in all parent directories and make include paths relative to its location if found")]
        public string FromParentDirWithFile { get; set; } = "build.root";

        public enum UseFileRelativePathMode
        {
            Always,                     // Always use including file's directory first to try to resolve include paths
            Never,                      // Never use including file's directory to try to resolve include paths
            OnlyInSameDirectory,        // Paths relative to the including file are only used if they resolve to the directory containing the including file (not including subdirectories)
            OnlyInSameOrSubDirectory    // Paths relative to the including file are only used if they resolve to the directory containing the including file (including subdirectories)
        }
        [Category("Path")]
        [DisplayName("Use File Relative Paths")]
        [Description("Whether and when to use include paths relative to the current file.")]
        public UseFileRelativePathMode UseFileRelativePath { get; set; } = UseFileRelativePathMode.Never;

        //[Category("Path")]
        //[DisplayName("Ignore Standard Library")]
        //[Description("")]
        //public bool IgnorePathForStdLib { get; set; } = true;

        #endregion

        #region Formatting

        public enum DelimiterMode
        {
            Unchanged,
            AngleBrackets,
            Quotes,
        }
        [Category("Formatting")]
        [DisplayName("Delimiter Mode")]
        [Description("Optionally changes all delimiters to either angle brackets <...> or quotes \"...\".")]
        public DelimiterMode DelimiterFormatting { get; set; } = DelimiterMode.Unchanged;

        public enum SlashMode
        {
            Unchanged,
            ForwardSlash,
            BackSlash,
        }
        [Category("Formatting")]
        [DisplayName("Slash Mode")]
        [Description("Changes all slashes to the given type.")]
        public SlashMode SlashFormatting { get; set; } = SlashMode.ForwardSlash;

        [Category("Formatting")]
        [DisplayName("Remove Empty Lines")]
        [Description("If true, all empty lines of a include selection will be removed.")]
        public bool RemoveEmptyLines { get; set; } = true;

        #endregion

        #region Sorting

        [Category("Sorting")]
        [DisplayName("Include delimiters in precedence regexes")]
        [Description("If true, precedence regexes will consider delimiters (angle brackets or quotes.)")]
        public bool RegexIncludeDelimiter { get; set; } = true;

        [Category("Sorting")]
        [DisplayName("Insert blank line between precedence regex match groups")]
        [Description("If true, a blank line will be inserted after each group matching one of the precedence regexes.")]
        public bool BlankAfterRegexGroupMatch { get; set; } = true;

        [Category("Sorting")]
        [DisplayName("Precedence Regexes")]
        [Description("Earlier match means higher sorting priority.\n\"" + RegexUtils.CurrentFileNameKey + "\" will be replaced with the current file name without extension.")]
        public string[] PrecedenceRegexes {
            get { return precedenceRegexes; }
            set { precedenceRegexes = value.Where(x => x.Length > 0).ToArray(); } // Remove empty lines.
        }
        private string[] precedenceRegexes = new string[] {
            "(?i)stdafx\\.h(?-i)",
            $"\"(?i){RegexUtils.CurrentFileNameKey}\\.(h|hpp|hxx)(?-i)\"",
            "\"[^\\/]+\"",
            "\"[^.\\/]+(\\\\|/)",
            "<[^.\\/]+(\\\\|/)",
            "<[^\\/]+\\.[^.\\/]+>",
            "<[^.\\/]+>"
        };

        public enum TypeSorting
        {
            None,
            AngleBracketsFirst,
            QuotedFirst,
        }
        [Category("Sorting")]
        [DisplayName("Sort by Include Type")]
        [Description("Optionally put either includes with angle brackets <...> or quotes \"...\" first.")]
        public TypeSorting SortByType { get; set; } = TypeSorting.None;

        [Category("Sorting")]
        [DisplayName("Remove duplicates")]
        [Description("If true, duplicate includes will be removed.")]
        public bool RemoveDuplicates { get; set; } = true;

        #endregion

        public override void SaveSettingsToStorage()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var settingsStore = GetSettingsStore();

            if (!settingsStore.CollectionExists(CollectionName))
                settingsStore.CreateCollection(CollectionName);

            settingsStore.SetInt32(CollectionName, nameof(PathFormat), (int)PathFormat);
            settingsStore.SetString(CollectionName, nameof(FromParentDirWithFile), FromParentDirWithFile);
            settingsStore.SetInt32(CollectionName, nameof(UseFileRelativePath), (int)UseFileRelativePath);

            settingsStore.SetInt32(CollectionName, nameof(DelimiterFormatting), (int)DelimiterFormatting);
            settingsStore.SetInt32(CollectionName, nameof(SlashFormatting), (int)SlashFormatting);
            settingsStore.SetBoolean(CollectionName, nameof(RemoveEmptyLines), RemoveEmptyLines);

            settingsStore.SetBoolean(CollectionName, nameof(RegexIncludeDelimiter), RegexIncludeDelimiter);
            settingsStore.SetBoolean(CollectionName, nameof(BlankAfterRegexGroupMatch), BlankAfterRegexGroupMatch);
            var value = string.Join("\n", PrecedenceRegexes);
            settingsStore.SetString(CollectionName, nameof(PrecedenceRegexes), value);
            settingsStore.SetInt32(CollectionName, nameof(SortByType), (int)SortByType);
            settingsStore.SetBoolean(CollectionName, nameof(RemoveDuplicates), RemoveDuplicates);
        }

        public override void LoadSettingsFromStorage()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var settingsStore = GetSettingsStore();

            if (settingsStore.PropertyExists(CollectionName, nameof(PathFormat)))
                PathFormat = (PathMode)settingsStore.GetInt32(CollectionName, nameof(PathFormat));
            if (settingsStore.PropertyExists(CollectionName, nameof(FromParentDirWithFile)))
                FromParentDirWithFile = settingsStore.GetString(CollectionName, nameof(FromParentDirWithFile));
            if (settingsStore.PropertyExists(CollectionName, nameof(UseFileRelativePath)))
                UseFileRelativePath = (UseFileRelativePathMode)settingsStore.GetInt32(CollectionName, nameof(UseFileRelativePath));

            if (settingsStore.PropertyExists(CollectionName, nameof(DelimiterFormatting)))
                DelimiterFormatting = (DelimiterMode) settingsStore.GetInt32(CollectionName, nameof(DelimiterFormatting));
            if (settingsStore.PropertyExists(CollectionName, nameof(SlashFormatting)))
                SlashFormatting = (SlashMode) settingsStore.GetInt32(CollectionName, nameof(SlashFormatting));
            if (settingsStore.PropertyExists(CollectionName, nameof(RemoveEmptyLines)))
                RemoveEmptyLines = settingsStore.GetBoolean(CollectionName, nameof(RemoveEmptyLines));

            if (settingsStore.PropertyExists(CollectionName, nameof(RegexIncludeDelimiter)))
                RegexIncludeDelimiter = settingsStore.GetBoolean(CollectionName, nameof(RegexIncludeDelimiter));
            if (settingsStore.PropertyExists(CollectionName, nameof(BlankAfterRegexGroupMatch)))
                BlankAfterRegexGroupMatch = settingsStore.GetBoolean(CollectionName, nameof(BlankAfterRegexGroupMatch));
            if (settingsStore.PropertyExists(CollectionName, nameof(PrecedenceRegexes)))
            {
                var value = settingsStore.GetString(CollectionName, nameof(PrecedenceRegexes));
                PrecedenceRegexes = value.Split(new[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
            }
            if (settingsStore.PropertyExists(CollectionName, nameof(SortByType)))
                SortByType = (TypeSorting) settingsStore.GetInt32(CollectionName, nameof(SortByType));
            if (settingsStore.PropertyExists(CollectionName, nameof(RemoveDuplicates)))
                RemoveDuplicates = settingsStore.GetBoolean(CollectionName, nameof(RemoveDuplicates));
        }
    }
}

using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;

namespace IncludeToolbox
{
    // Notes:
    // Include Type
    //     General
    //     Precompiled Header
    //     Project Includes
    //     External Includes
    //     System Includes

    [Guid("A220A8FD-2C84-41BD-BF7C-97C30E9CB5DB")]
    public class IncludeTypeOptionsPage : OptionsPage
    {
        public const string SubCategory = FormatterOptionsPage.SubCategory + @"\Include Type";
        public const string CollectionName = FormatterOptionsPage.CollectionName + @"\IncludeType";

        [Category("Format by Include Type")]
        [DisplayName("Enable")]
        [Description("Enable formatting by include type.")]
        public bool FormatByIncludeType { get; set; } = false;

        public override void SaveSettingsToStorage()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var settingsStore = GetSettingsStore();

            if (!settingsStore.CollectionExists(CollectionName))
                settingsStore.CreateCollection(CollectionName);

            settingsStore.SetBoolean(CollectionName, nameof(FormatByIncludeType), FormatByIncludeType);
        }

        public override void LoadSettingsFromStorage()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var settingsStore = GetSettingsStore();

            if (settingsStore.PropertyExists(CollectionName, nameof(FormatByIncludeType)))
                FormatByIncludeType = settingsStore.GetBoolean(CollectionName, nameof(FormatByIncludeType));
        }
    }
}

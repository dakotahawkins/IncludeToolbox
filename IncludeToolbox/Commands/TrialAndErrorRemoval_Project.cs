using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;

namespace IncludeToolbox.Commands
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class TrialAndErrorRemoval_Project : CommandBase<TrialAndErrorRemoval_Project>
    {
        public override CommandID CommandID => new CommandID(CommandSetGuids.ProjectGroup, 0x0100);

        private TrialAndErrorRemoval impl;
        private TrialAndErrorRemovalOptionsPage settings;

        private ProjectItems projectItems = null;
        private int numTotalRemovedIncludes = 0;
        private Queue<ProjectItem> projectFiles;

        public TrialAndErrorRemoval_Project()
        {
            projectFiles = new Queue<ProjectItem>();
        }

        protected override void SetupMenuCommand()
        {
            base.SetupMenuCommand();

            impl = new TrialAndErrorRemoval();
            impl.OnFileFinished += OnDocumentIncludeRemovalFinished;
            menuCommand.BeforeQueryStatus += UpdateVisibility;

            settings = (TrialAndErrorRemovalOptionsPage)Package.GetDialogPage(typeof(TrialAndErrorRemovalOptionsPage));
        }

        private void OnDocumentIncludeRemovalFinished(int removedIncludes, bool canceled)
        {
            numTotalRemovedIncludes += removedIncludes;
            if (canceled || !ProcessNextFile())
            {
                Output.Instance.InfoMsg("Removed total of {0} #include directives from project.", numTotalRemovedIncludes);
                numTotalRemovedIncludes = 0;
            }
        }

        private void UpdateVisibility(object sender, EventArgs e)
        {
            string reason;
            var project = GetSelectedCppProject(out reason);
            menuCommand.Visible = project != null;
        }

        static Project GetSelectedCppProject(out string reasonForFailure)
        {
            reasonForFailure = "";

            var selectedItems = VSUtils.GetDTE().SelectedItems;
            if (selectedItems.Count < 1)
            {
                reasonForFailure = "Selection is empty!";
                return null;
            }

            // Reading .Item(object) behaves weird, but iterating works.
            foreach (SelectedItem item in selectedItems)
            {
                Project vcProject = item?.Project;
                if (VSUtils.VCUtils.IsVCProject(vcProject))
                {
                    return vcProject;
                }
            }

            reasonForFailure = "Selection does not contain a C++ project!";
            return null;
        }

        private bool ProcessNextFile()
        {
            while (projectFiles.Count > 0)
            {
                ProjectItem projectItem = projectFiles.Dequeue();

                Document document = null;
                try
                {
                    document = projectItem.Open().Document;
                }
                catch (Exception)
                {
                }
                if (document == null)
                    continue;

                bool started = impl.PerformTrialAndErrorIncludeRemoval(document, settings);
                if (started)
                    return true;
            }
            return false;
        }

        static void RecursiveFindFilesInProject(ProjectItems items, ref Queue<ProjectItem> projectFiles)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            var e = items.GetEnumerator();
            while (e.MoveNext())
            {
                var item = e.Current;
                if (item == null)
                    continue;
                var projectItem = item as ProjectItem;
                if (projectItem == null)
                    continue;

                Guid projectItemKind = new Guid(projectItem.Kind);
                if (projectItemKind == VSConstants.GUID_ItemType_VirtualFolder ||
                    projectItemKind == VSConstants.GUID_ItemType_PhysicalFolder)
                {
                    RecursiveFindFilesInProject(projectItem.ProjectItems, ref projectFiles);
                }
                else if (projectItemKind == VSConstants.GUID_ItemType_PhysicalFile)
                {
                    projectFiles.Enqueue(projectItem);
                }
                else
                {
                    Output.Instance.WriteLine("Unexpected Error: Unknown projectItem {0} of Kind {1}", projectItem.Name, projectItem.Kind);
                }
            }
        }

        private void PerformTrialAndErrorRemoval(Project project)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            projectItems = project.ProjectItems;

            projectFiles.Clear();
            RecursiveFindFilesInProject(projectItems, ref projectFiles);

            if (projectFiles.Count > 2)
            {
                if (Output.Instance.YesNoMsg("Attention! Trial and error include removal on large projects make take up to several hours! In this time you will not be able to use Visual Studio. Are you sure you want to continue?")
                    != Output.MessageResult.Yes)
                {
                    return;
                }
            }

            numTotalRemovedIncludes = 0;
            ProcessNextFile();
        }
    

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        protected override void MenuItemCallback(object sender, EventArgs e)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            if (TrialAndErrorRemoval.WorkInProgress)
            {
                Output.Instance.ErrorMsg("Trial and error include removal already in progress!");
                return;
            }

            try
            {
                Project project = GetSelectedCppProject(out string reasonForFailure);
                if (project == null)
                {
                    Output.Instance.WriteLine(reasonForFailure);
                    return;
                }

                PerformTrialAndErrorRemoval(project);
            }
            finally
            {
                projectItems = null;
            }
        }
    }
}

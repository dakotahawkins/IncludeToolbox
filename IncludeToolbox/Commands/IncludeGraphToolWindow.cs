using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace IncludeToolbox.Commands
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class IncludeGraphToolWindow : CommandBase<IncludeGraphToolWindow>
    {
        public override CommandID CommandID => new CommandID(CommandSetGuids.ToolGroup, 0x0102);

        /// <summary>
        /// Shows the tool window when the menu item is clicked.
        /// </summary>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event args.</param>
        protected override void MenuItemCallback(object sender, EventArgs e)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            // Get the instance number 0 of this tool window. This window is single instance so this instance
            // is actually the only one.
            // The last flag is set to true so that if the tool window does not exists it will be created.
            ToolWindowPane window = Package.FindToolWindow(typeof(GraphWindow.IncludeGraphToolWindow), 0, true);
            if (window?.Frame == null)
            {
                Output.Instance.ErrorMsg("Failed to open Include Graph window!");
            }
            else
            {
                Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
                IVsWindowFrame windowFrame = (IVsWindowFrame)window.Frame;
                windowFrame.SetProperty((int)__VSFPROPID.VSFPROPID_CmdUIGuid, GraphWindow.IncludeGraphToolWindow.GUIDString);
                Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(windowFrame.Show());
            }
        }
    }
}

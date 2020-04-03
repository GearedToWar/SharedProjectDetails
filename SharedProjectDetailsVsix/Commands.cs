using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace SharedProjectDetailsVsix
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class Commands
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int GetSharedProjectsCommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("408497a5-714e-45fd-a067-b7dfe79c4a15");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        private const string SharedProjectDetailsWindowPaneCaption = "Shared project details";
        private IVsOutputWindowPane windowPane;

        /// <summary>
        /// Initializes a new instance of the <see cref="Commands"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private Commands(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, GetSharedProjectsCommandId);
            var menuItem = new MenuCommand(this.GetSharedProjectDetails, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static Commands Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in Command's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new Commands(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "VSTHRD102:Implement internal logic asynchronously", Justification = "Call-back for a command")]
        private void GetSharedProjectDetails(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            package.JoinableTaskFactory.Run(async () =>
            {
                var solution = await this.package.GetServiceAsync<SVsSolution, IVsSolution>();

                EnumerateSharedProjectsInSolution(solution);
            });
        }

        private void EnumerateSharedProjectsInSolution(IVsSolution solution)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            int hr = solution.GetProjectEnum((uint)__VSENUMPROJFLAGS.EPF_LOADEDINSOLUTION, Guid.Empty, out IEnumHierarchies enumHierarchies);

            if (ErrorHandler.Failed(hr))
            {
                return;
            }

            IVsHierarchy[] hierarchies = new IVsHierarchy[1];

            while (ErrorHandler.Succeeded(enumHierarchies.Next(1, hierarchies, out uint fetchedCount)) && (fetchedCount == 1))
            {
                if (hierarchies[0].IsSharedAssetsProject())
                {
                    hierarchies[0].GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_Name, out object result);
                    WriteToOutputPane("Shared project: " + result as string);

                    WriteToOutputPane("\tProject items: ");
                    foreach (var itemId in GetProjectItems(hierarchies[0], (uint)VSConstants.VSITEMID.Root))
                    {
                        hr = hierarchies[0].GetCanonicalName(itemId, out string projectItemName);
                        WriteToOutputPane($"\t\tName: [{projectItemName}] Item id: [{(ErrorHandler.Succeeded(hr) ? itemId.ToString() : "NULL")}]");
                    }

                    WriteToOutputPane("Importing projects: ");

                    foreach (var importingProject in hierarchies[0].EnumImportingProjects())
                    {
                        importingProject.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_Name, out result);
                        WriteToOutputPane("\t" + result as string);

                        WriteToOutputPane("\tShared Project items: ");
                        foreach (var itemId in GetProjectItems(importingProject, (uint)VSConstants.VSITEMID.Root))
                        {
                            if (SharedProjectUtilities.IsSharedItem(importingProject, itemId))
                            {
                                hr = importingProject.GetCanonicalName(itemId, out string projectItemName);
                                WriteToOutputPane($"\t\tName: [{projectItemName}] Item id: [{(ErrorHandler.Succeeded(hr) ? itemId.ToString() : "NULL")}]");
                            }
                        }
                    }

                    hierarchies[0].GetActiveProjectContext().GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_Name, out result);
                    WriteToOutputPane("Active project context = " + result as string);
                }
            }
        }

        internal IEnumerable<uint> GetProjectItems(IVsHierarchy projectHierarchy, uint itemId)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            int hr = projectHierarchy.GetProperty(itemId, (int)__VSHPROPID.VSHPROPID_FirstChild, out object childItemIdObject);

            while (ErrorHandler.Succeeded(hr))
            {
                if (uint.TryParse(childItemIdObject.ToString(), out uint childItemId))
                {
                    // Read the canonical name of the child item in the project hierarchy
                    // hr = projectHierarchy.GetCanonicalName(childItemId, out string projectItemName);

                    yield return childItemId;

                    foreach (uint descendantItemId in GetProjectItems(projectHierarchy, childItemId))
                    {
                        yield return descendantItemId;
                    }

                    hr = projectHierarchy.GetProperty(childItemId, (int)__VSHPROPID.VSHPROPID_NextSibling, out childItemIdObject);
                }
                else
                {
                    hr = VSConstants.E_FAIL;
                }
            }
        }

        private void WriteToOutputPane(string data)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (this.windowPane == null)
            {
                this.windowPane = this.package.GetOutputPane(VSConstants.OutputWindowPaneGuid.GeneralPane_guid, SharedProjectDetailsWindowPaneCaption);
            }

            int? hr = this.windowPane?.Activate();

            if (hr.HasValue && ErrorHandler.Succeeded(hr.Value))
            {
                this.windowPane?.OutputStringThreadSafe(data + Environment.NewLine);
            }
        }
    }
}

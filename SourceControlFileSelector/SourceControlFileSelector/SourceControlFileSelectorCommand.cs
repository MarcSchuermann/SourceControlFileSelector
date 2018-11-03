﻿//// --------------------------------------------------------------------------------------------------------------------
//// <copyright>Marc Schürmann</copyright>
//// --------------------------------------------------------------------------------------------------------------------

using Microsoft.VisualStudio.Shell;
using SourceControlFileSelector.tfsAccess;
using System;
using System.ComponentModel.Design;
using Task = System.Threading.Tasks.Task;

namespace SourceControlFileSelector
{
    /// <summary>Command handler</summary>
    internal sealed class SourceControlFileSelectorCommand
    {
        #region Public Fields

        /// <summary>Command ID.</summary>
        public const int CommandId = 0x0100;

        /// <summary>Command menu group (command set GUID).</summary>
        public static readonly Guid CommandSet = new Guid("50ce47ea-ed9d-49b7-87f3-e31d9ee84f35");

        #endregion Public Fields

        #region Private Fields

        /// <summary>VS Package that provides this command, not null.</summary>
        private readonly AsyncPackage package;

        #endregion Private Fields

        #region Private Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="SourceControlFileSelectorCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private SourceControlFileSelectorCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        #endregion Private Constructors

        #region Public Properties

        /// <summary>Gets the instance of the command.</summary>
        public static SourceControlFileSelectorCommand Instance
        {
            get;
            private set;
        }

        #endregion Public Properties

        #region Private Properties

        private static EnvDTE.DTE dte { get; set; }

        /// <summary>Gets the service provider from the owner package.</summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        #endregion Private Properties

        #region Public Methods

        /// <summary>Initializes the singleton instance of the command.</summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in
            // SourceControlFileSelectorCommand's constructor requires the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new SourceControlFileSelectorCommand(package, commandService);
            dte = await package.GetServiceAsync(typeof(EnvDTE.DTE)) as EnvDTE.DTE;
        }

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var selectedItem = GetFirstItem(dte?.SelectedItems);
            if (selectedItem != null)
            {
                var localPath = GetLocalePath(selectedItem);
                var tfs = new TfsWrapper();

                if (tfs != null)
                {
                    var versionControlServer = tfs.GetVersionControlServer();
                    var workspace = versionControlServer.GetWorkspace(localPath);
                    var serverPath = workspace.TryGetServerItemForLocalItem(localPath);

                    var sourceControlExplorer = tfs.GetSourceControlExplorer();
                    tfs.SelectInSourceControlExplorer(serverPath, workspace);
                }
            }
        }

        private EnvDTE.SelectedItem GetFirstItem(EnvDTE.SelectedItems selectedItems)
        {
            if (selectedItems == null)
                return null;

            foreach (EnvDTE.SelectedItem item in selectedItems)
            {
                return item;
            }

            return null;
        }

        private string GetLocalePath(EnvDTE.SelectedItem item)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string result = string.Empty;

            if (item.ProjectItem == null)
            {
                if (item.Project == null)
                {
                    result = dte.Solution.FullName;
                }
                else
                {
                    result = item.Project.FullName;
                }
            }
            else
            {
                try
                {
                    result = item.ProjectItem.get_FileNames(0);
                }
                catch (ArgumentException)
                {
                    result = item.ProjectItem.get_FileNames(1);
                }
            }
            return result;
        }

        #endregion Private Methods
    }
}
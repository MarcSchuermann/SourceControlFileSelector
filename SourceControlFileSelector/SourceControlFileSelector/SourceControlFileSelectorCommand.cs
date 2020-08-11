//// --------------------------------------------------------------------------------------------------------------------
//// <copyright>Marc Schürmann</copyright>
//// --------------------------------------------------------------------------------------------------------------------

using EnvDTE;
using Microsoft.VisualStudio.Shell;
using SourceControlFileSelector.Misc;
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

        private static DTE dte { get; set; }

        private static EnvDTE80.DTE2 dte2 { get; set; }

        /// <summary>Gets the service provider from the owner package.</summary>
        private IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return package;
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

            dte = await package.GetServiceAsync(typeof(DTE)) as DTE;
            dte2 = await package.GetServiceAsync(typeof(DTE)) as EnvDTE80.DTE2;
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
            var logger = new TraceLogger(dte);

            if (dte2?.ActiveWindow != null && dte2?.ActiveWindow == dte2?.ActiveDocument?.ActiveWindow)
            {
              var activeDocumentLocalPath = dte2?.ActiveDocument?.FullName;
              if (!string.IsNullOrWhiteSpace(activeDocumentLocalPath))
              {
                SelectLocalPath(activeDocumentLocalPath, logger, dte2.ActiveDocument.Name);
                return;
              }
            }

            if (dte2?.SelectedItems?.Count > 0)
            {
                foreach (SelectedItem selectedItem in dte2.SelectedItems)
                {
                    var localPath = GetLocalPath(selectedItem);
                    if (!string.IsNullOrWhiteSpace(localPath))
                    {
                        logger.Log($"The selected item is '{selectedItem.Name}'.");
                        SelectLocalPath(localPath, logger, selectedItem.Name);
                        return;
                    }
                }
            }

            logger.Log($"No file with local path selected.");
        }

        private void SelectLocalPath(string localPath, TraceLogger logger, string name)
        {
            logger.Log($"The local path is '{localPath}'.");

            var tfs = new TfsWrapper();
            logger.Log($"The tfs is '{tfs}'.");

            var versionControlServer = tfs.GetVersionControlServer();
            if (versionControlServer != null)
            {
                logger.Log($"The versionControlServer is '{versionControlServer}'.");

                var workspace = versionControlServer.GetWorkspace(localPath);
                logger.Log($"The workspace is '{workspace.Name}'.");
                var serverPath = workspace.TryGetServerItemForLocalItem(localPath);
                logger.Log($"The serverPath is '{serverPath}'.");

                var sourceControlExplorer = tfs.GetSourceControlExplorer();
                tfs.SelectInSourceControlExplorer(serverPath, workspace, sourceControlExplorer);
            }
            else
            {
                logger.Log($"The file {localPath} is not under source control.");
            }

            logger.Log($"End of selecting '{name}'.");
        }

        private string GetLocalPath(SelectedItem item)
        {
            if (item.ProjectItem != null)
            {
                try
                {
                    return item.ProjectItem.get_FileNames(0);
                }
                catch(ArgumentException)
                {
                    return item.ProjectItem.get_FileNames(1);
                }
            }

            return item.Project?.FullName ?? dte2?.Solution?.FullName;
        }

        #endregion Private Methods
    }
}
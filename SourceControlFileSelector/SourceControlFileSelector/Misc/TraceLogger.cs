//// --------------------------------------------------------------------------------------------------------------------
//// <copyright>Marc Schürmann</copyright>
//// --------------------------------------------------------------------------------------------------------------------

using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using System;

namespace SourceControlFileSelector.Misc
{
    /// <summary>The Trace logger.</summary>
    public class TraceLogger
    {
        #region Private Fields

        private DTE2 dte2;
        private OutputWindowPane outputPane;

        #endregion Private Fields

        #region Public Constructors

        /// <summary>The trace logger.</summary>
        /// <param name="dte"></param>
        public TraceLogger(DTE2 dte2)
        {
            this.dte2 = dte2;
            outputPane = GetOutputPane();

            outputPane.Activate();
            Log("Successfuly created source control file selector logger");
        }

        #endregion Public Constructors

        #region Public Methods

        /// <summary>Logs the specified message.</summary>
        /// <param name="message">The message.</param>
        public void Log(string message)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            outputPane?.Activate();
            outputPane?.OutputString(Environment.NewLine);
            outputPane?.OutputString(message);
        }

        #endregion Public Methods

        #region Private Methods

        /// <summary>Gets the output pane.</summary>
        /// <returns>The output pane.</returns>
        private OutputWindowPane GetOutputPane()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (outputPane == null)
            {
                var outputWindow = dte2.ToolWindows.OutputWindow;
                outputPane = outputWindow.OutputWindowPanes.Add("Source control file selector");
            }

            return outputPane;
        }

        #endregion Private Methods
    }
}
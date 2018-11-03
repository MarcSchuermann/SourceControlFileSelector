//// --------------------------------------------------------------------------------------------------------------------
//// <copyright>Marc Schürmann</copyright>
//// --------------------------------------------------------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.Reflection;

namespace SourceControlFileSelector.tfsAccess
{
    /// <summary>The team foundation server wrapper.</summary>
    public class TfsWrapper
    {
        #region Private Fields

        private dynamic hatterasService = null;
        private dynamic wrapped = null;

        #endregion Private Fields

        #region Public Constructors

        public TfsWrapper()
        {
            try
            {
                var _vcAssembly = Assembly.Load("Microsoft.VisualStudio.TeamFoundation.VersionControl");
                Type t = _vcAssembly.GetType("Microsoft.VisualStudio.TeamFoundation.VersionControl.HatPackage");
                var prop = t.GetProperty("Instance", BindingFlags.NonPublic | BindingFlags.Static);
                wrapped = new AccessPrivateWrapper(prop.GetValue(null, null));
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        #endregion Public Constructors

        #region Private Properties

        private dynamic HatterasService
        {
            get
            {
                if (hatterasService == null)
                {
                    hatterasService = new AccessPrivateWrapper(wrapped.HatterasService);
                }
                return hatterasService;
            }
            set { hatterasService = value; }
        }

        #endregion Private Properties

        #region Public Methods

        public object GetSourceControlExplorer()
        {
            dynamic sccToolWindow = new AccessPrivateWrapper(wrapped.GetToolWindowSccExplorer(true));
            dynamic explorer = new AccessPrivateWrapper(sccToolWindow.SccExplorer);

            return explorer;
        }

        public dynamic GetVersionControlServer()
        {
            return HatterasService.VersionControlServer;
        }

        #endregion Public Methods

        #region Internal Methods

        internal void SelectInSourceControlExplorer(dynamic serverPath, dynamic workspace)
        {
            wrapped.OpenSceToPath(serverPath, workspace);
        }

        #endregion Internal Methods
    }
}
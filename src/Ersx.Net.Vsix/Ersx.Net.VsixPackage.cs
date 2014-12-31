using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Win32;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;

namespace seldary.Ersx_Net_Vsix
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the 
    /// IVsPackage interface and uses the registration attributes defined in the framework to 
    /// register itself and its components with the shell.
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidErsx_Net_VsixPkgString)]
    [ProvideAutoLoad(UIContextGuids80.SolutionExists)]
    public sealed class Ersx_Net_VsixPackage : Package
    {
        #region Package Members

        protected override void Initialize()
        {
            base.Initialize();

            var oleMenuCommandService = GetService(typeof (IMenuCommandService)) as OleMenuCommandService;
            if (oleMenuCommandService != null)
            {
                var menuCommandId = new CommandID(GuidList.guidErsx_Net_VsixCmdSet, (int) PkgCmdIDList.sortResx);
                var menuItem = new OleMenuCommand(MenuItemClick, menuCommandId);
                menuItem.BeforeQueryStatus += BeforeContextMenuOpens;
                oleMenuCommandService.AddCommand(menuItem);
            }
        }

        private void BeforeContextMenuOpens(object sender, EventArgs e)
        {
            var menuCommand = sender as OleMenuCommand;
            if (menuCommand == null)
            {
                return;
            }

            IVsHierarchy hierarchy;
            uint itemId;
            if (IsSingleProjectItemSelection(out hierarchy, out itemId))
            {
                string itemFullPath;
                ((IVsProject) hierarchy).GetMkDocument(itemId, out itemFullPath);

                var isResxFile = new FileInfo(itemFullPath).Extension.Equals(".resx", StringComparison.OrdinalIgnoreCase);
                ToggleMenuCommand(menuCommand, isResxFile);
            }
            else
            {
                ToggleMenuCommand(menuCommand, false);
            }
        }

        private void ToggleMenuCommand(OleMenuCommand menuCommand, bool isEnabled)
        {
            menuCommand.Enabled = isEnabled;
            menuCommand.Visible = isEnabled;
        }

        private bool IsSingleProjectItemSelection(out IVsHierarchy hierarchy, out uint itemId)
        {
            hierarchy = null;
            itemId = VSConstants.VSITEMID_NIL;

            var monitorSelection = GetGlobalService(typeof(SVsShellMonitorSelection)) as IVsMonitorSelection;
            var solution = GetGlobalService(typeof(SVsSolution)) as IVsSolution;
            if (monitorSelection == null || solution == null)
            {
                return false;
            }

            var hierarchyPtr = IntPtr.Zero;
            var selectionContainerPtr = IntPtr.Zero;

            try
            {
                IVsMultiItemSelect multiItemSelect;
                var currentSelectionResult = monitorSelection.GetCurrentSelection(out hierarchyPtr, out itemId, out multiItemSelect, out selectionContainerPtr);
                if (ErrorHandler.Failed(currentSelectionResult) || 
                    hierarchyPtr == IntPtr.Zero || 
                    itemId == VSConstants.VSITEMID_NIL || 
                    multiItemSelect != null || 
                    itemId == VSConstants.VSITEMID_ROOT)
                {
                    return false;
                }

                hierarchy = Marshal.GetObjectForIUnknown(hierarchyPtr) as IVsHierarchy;
                if (hierarchy == null)
                {
                    return false;
                }

                Guid guidProjectId;
                return !ErrorHandler.Failed(solution.GetGuidOfProject(hierarchy, out guidProjectId));
            }
            finally
            {
                if (selectionContainerPtr != IntPtr.Zero)
                {
                    Marshal.Release(selectionContainerPtr);
                }

                if (hierarchyPtr != IntPtr.Zero)
                {
                    Marshal.Release(hierarchyPtr);
                }
            }
        }

        #endregion

        private void MenuItemClick(object sender, EventArgs e)
        {
            var menuCommand = sender as OleMenuCommand;
            if (menuCommand != null)
            {
                IVsHierarchy hierarchy;
                uint itemId;
                if (!IsSingleProjectItemSelection(out hierarchy, out itemId))
                {
                    return;
                }
                string itemFullPath;
                ((IVsProject)hierarchy).GetMkDocument(itemId, out itemFullPath);
                var fileInfo = new FileInfo(itemFullPath);
                XDocument doc;
                using (var streamReader = fileInfo.OpenText())
                {
                    doc = XDocument.Load(streamReader);
                }

                using (var fileStream = fileInfo.OpenWrite())
                {
                    var sortedDoc = SortDataByName(doc);
                    sortedDoc.Save(fileStream);
                }
            }
        }

        private static XDocument SortDataByName(XDocument resx)
        {
            return new XDocument(
                new XElement(resx.Root.Name,
                    from comment in resx.Root.Nodes() where comment.NodeType == XmlNodeType.Comment select comment,
                    from schema in resx.Root.Elements() where schema.Name.LocalName == "schema" select schema,
                    from resheader in resx.Root.Elements("resheader") orderby (string)resheader.Attribute("name") select resheader,
                    from assembly in resx.Root.Elements("assembly") orderby (string)assembly.Attribute("name") select assembly,
                    from metadata in resx.Root.Elements("metadata") orderby (string)metadata.Attribute("name") select metadata,
                    from data in resx.Root.Elements("data") orderby (string)data.Attribute("name") select data
                )
            );
        }
    }
}
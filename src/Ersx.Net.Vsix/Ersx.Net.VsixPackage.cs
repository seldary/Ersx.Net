using System;
using System.ComponentModel.Design;
using System.IO;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using Ersx.Net;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace seldary.Ersx_Net_Vsix
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidErsx_Net_VsixPkgString)]
    [ProvideAutoLoad(UIContextGuids80.SolutionExists)]
    public sealed class Ersx_Net_VsixPackage : Package
    {
        private readonly ResxSorter _resxSorter = new ResxSorter();

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

            var monitorSelection = GetGlobalService(typeof (SVsShellMonitorSelection)) as IVsMonitorSelection;
            var solution = GetGlobalService(typeof (SVsSolution)) as IVsSolution;
            if (monitorSelection == null || solution == null)
            {
                return false;
            }

            var hierarchyPtr = IntPtr.Zero;
            var selectionContainerPtr = IntPtr.Zero;

            try
            {
                IVsMultiItemSelect multiItemSelect;
                var currentSelectionResult = monitorSelection.GetCurrentSelection(out hierarchyPtr, out itemId,
                    out multiItemSelect, out selectionContainerPtr);
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
            if (menuCommand == null)
            {
                return;
            }

            IVsHierarchy hierarchy;
            uint itemId;
            if (!IsSingleProjectItemSelection(out hierarchy, out itemId))
            {
                return;
            }
            string itemFullPath;
            ((IVsProject) hierarchy).GetMkDocument(itemId, out itemFullPath);
            _resxSorter.
                Sort(XDocument.Load(itemFullPath)).
                Save(itemFullPath);
        }
    }
}
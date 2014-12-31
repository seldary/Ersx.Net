using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
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
            var isResxFileSelected = GetSelectedResxFilePaths().Any();
            menuCommand.Enabled = isResxFileSelected;
            menuCommand.Visible = isResxFileSelected;
        }

        private IEnumerable<string> GetSelectedFileFullPaths()
        {
            IVsMultiItemSelect multiItemSelect;
            uint itemId;
            IntPtr hierarchyPtr;
            IntPtr selectionContainerPtr;
            var currentSelectionResult = (GetGlobalService(typeof (SVsShellMonitorSelection)) as IVsMonitorSelection).
                GetCurrentSelection(out hierarchyPtr, out itemId, out multiItemSelect, out selectionContainerPtr);
            if (ErrorHandler.Failed(currentSelectionResult) ||
                hierarchyPtr == IntPtr.Zero ||
                itemId == VSConstants.VSITEMID_NIL ||
                itemId == VSConstants.VSITEMID_ROOT)
            {
                return Enumerable.Empty<string>();
            }

            if (multiItemSelect == null)
            {
                return new[]
                {
                    GetSelectedFileFullPath(itemId, Marshal.GetObjectForIUnknown(hierarchyPtr) as IVsHierarchy)
                };
            }

            uint selectedItemCount;
            int pfSingleHirearchy;
            multiItemSelect.GetSelectionInfo(out selectedItemCount, out pfSingleHirearchy);
            var selectedItems = new VSITEMSELECTION[selectedItemCount];
            multiItemSelect.GetSelectedItems(0, selectedItemCount, selectedItems);
            return selectedItems.Select(_ => GetSelectedFileFullPath(_.itemid, _.pHier));
        }

        private string GetSelectedFileFullPath(uint itemid, IVsHierarchy hierarchy)
        {
            string filePath;
            hierarchy.GetCanonicalName(itemid, out filePath);
            return filePath;
        }

        private IEnumerable<string> GetSelectedResxFilePaths()
        {
            return GetSelectedFileFullPaths().Where(IsResxFile);
        }

        private bool IsResxFile(string fullPath)
        {
            return fullPath.EndsWith(".resx", StringComparison.OrdinalIgnoreCase);
        }

        #endregion

        private void MenuItemClick(object sender, EventArgs e)
        {
            foreach (var selectedFileFullPath in GetSelectedResxFilePaths())
            {
                _resxSorter.
                    Sort(XDocument.Load(selectedFileFullPath)).
                    Save(selectedFileFullPath);
            }
        }
    }
}
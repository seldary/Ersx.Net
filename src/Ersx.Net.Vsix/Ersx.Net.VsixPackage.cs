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
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExistsAndFullyLoaded_string)]
    public sealed class Ersx_Net_VsixPackage : Package
    {
        private readonly ResxSorter _resxSorter = new ResxSorter();
        private Guid _resxOutputWindowId = new Guid("64FD8E73-DD6E-4697-A20D-B2307205A764");
        private IVsStatusbar _statusBar;
        private OleMenuCommandService _oleMenuCommandService;
        private IVsOutputWindow _outputWindow;
        private object _sortIcon;

        protected override void Initialize()
        {
            base.Initialize();

            _statusBar = GetService(typeof (SVsStatusbar)) as IVsStatusbar;
            _oleMenuCommandService = GetService(typeof (IMenuCommandService)) as OleMenuCommandService;
            _outputWindow = GetService(typeof (SVsOutputWindow)) as IVsOutputWindow;
            _sortIcon = (short) Constants.SBAI_General;

            if (_oleMenuCommandService != null)
            {
                var menuCommandId = new CommandID(GuidList.guidErsx_Net_VsixCmdSet, (int) PkgCmdIDList.sortResx);
                var menuItem = new OleMenuCommand(MenuItemClick, StatusChanged, BeforeContextMenuOpens, menuCommandId);
                _oleMenuCommandService.AddCommand(menuItem);
            }
        }

        private void StatusChanged(object sender, EventArgs e)
        {
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
            return GetSelectedFileFullPaths().Where(fullPath =>
                fullPath.EndsWith(".resx", StringComparison.OrdinalIgnoreCase));
        }

        private void MenuItemClick(object sender, EventArgs e)
        {
            try
            {
                uint cookie = 0;
                _statusBar.Progress(ref cookie, 1, "", 0, 0);
                _statusBar.Animation(1, ref _sortIcon);
                var selectedResxFilePaths = GetSelectedResxFilePaths().ToList();
                for (var i = 0; i < selectedResxFilePaths.Count; i++)
                {
                    _resxSorter.
                        Sort(XDocument.Load(selectedResxFilePaths[i])).
                        Save(selectedResxFilePaths[i]);

                    _statusBar.Progress(ref cookie, 1, "", (uint) i + 1, (uint) selectedResxFilePaths.Count);
                    _statusBar.SetText(string.Format("Sorting {0} out of {1} resx files...", i + 1,
                        selectedResxFilePaths.Count));
                }

                WriteToOutput(string.Join(Environment.NewLine,
                    new[] {string.Format("Sorted {0} resx files:", selectedResxFilePaths.Count())}.Concat(
                        selectedResxFilePaths)));
                _statusBar.Progress(ref cookie, 0, "", 0, 0);
                _statusBar.Animation(0, ref _sortIcon);
                SetStatusBar("Resx sort succeeded");
            }
            catch (Exception exception)
            {
                WriteToOutput("Failed to sort resx files: {0}", exception.Message);
                SetStatusBar("Failed to sort resx files");
            }
        }

        private void WriteToOutput(string message, params object[] messageParams)
        {
            _outputWindow.CreatePane(ref _resxOutputWindowId, "Resx", 1, 1);
            IVsOutputWindowPane customPane;
            _outputWindow.GetPane(ref _resxOutputWindowId, out customPane);
            customPane.Clear();
            customPane.OutputString(string.Format(message, messageParams));
            customPane.Activate();
        }

        private void SetStatusBar(string message, params object[] messageParams)
        {
            _statusBar.SetText(string.Format(message, messageParams));
        }
    }
}
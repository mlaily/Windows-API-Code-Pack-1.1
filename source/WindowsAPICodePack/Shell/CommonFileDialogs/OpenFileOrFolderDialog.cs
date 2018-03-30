//#define DEBUG_OFOFD

using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.WindowsAPICodePack.Controls;
using Microsoft.WindowsAPICodePack.Dialogs.Controls;
using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Shell.Interop.Common;
using MS.WindowsAPICodePack.Internal;

namespace Microsoft.WindowsAPICodePack.Dialogs
{
    /// <summary>
    /// A Windows 7 common file dialog, hacked to allow the user to select multiple files and folders at the same time.
    /// </summary>
    public sealed class OpenFileOrFolderDialog : CommonFileDialog
    {
        private const string CustomValidationButtonTextId = "OFOFD_CUSTOM_VALIDATION_BUTTON";
        private const string OkButtonTextId = "OFOFD_OK_BUTTON";
        private const string CancelButtonTextId = "OFOFD_CANCEL_BUTTON";
        private const string FileNameLabelTextId = "OFOFD_FILENAME_LABEL";

        private readonly string _okButtonUserText;
        private readonly string _cancelButtonUserText;

        private readonly CommonFileDialogButton _customValidationButton;

        private NativeFileOpenDialog _openDialogCoClass;

        // Create an instance of the delegate that won't be garbage collected to avoid:
        //   Managed Debugging Assistant 'CallbackOnCollectedDelegate' :**
        //   'A callback was made on a garbage collected delegate of type
        //   'WpfApp1!WpfApp1.MainWindow+NativeMethods+CBTProc::Invoke'.
        //   This may cause application crashes, corruption and data loss.
        //   When passing delegates to unmanaged code, they must be
        //   kept alive by the managed application until it is guaranteed
        //   that they will never be called.'
        private CallWndRetProc _callback;

        private ICollection<ShellObject> _userFilteredSelection;

        /// <summary>
        /// This event is called when the user tries to validate a selection.
        /// It allows to set a selection as invalid, and replace the selected files and folders with filtered values.
        /// </summary>
        public event EventHandler<ValidateSelectionEventArgs> ValidateSelection;

        /// <summary>
        /// A message shown when the user tries to validate the dialog with an invalid selection.
        /// The default dialog value will be used if null is provided.
        /// </summary>
        public string InvalidSelectionMessage { get; set; }
        private const string DefaultInvalidSelectionMessage = "Invalid selection!";

        /// <summary>
        /// Creates a new instance of this class.
        /// </summary>
        /// <param name="dialogTitle">The dialog title. The default dialog value will be used if null is provided.</param>
        /// <param name="okButtonText">Cannot be null.</param>
        /// <param name="cancelButtonText">The default dialog value will be used if null is provided.</param>
        /// <param name="invalidSelectionMessage">
        /// A message shown when the user tries to validate the dialog with an invalid selection.
        /// The default dialog value will be used if null is provided.
        /// </param>
        /// <param name="useSaneSelectionValidation">
        /// Subscribe <see cref="ValidateSelection"/> to a default implementation that ensures
        /// special/invalid folders and files cannot be selected, and .lnk shortcuts are dereferenced.
        /// </param>
        public OpenFileOrFolderDialog(
            string dialogTitle,
            string okButtonText,
            string cancelButtonText,
            string invalidSelectionMessage = null,
            bool useSaneSelectionValidation = true)
            : base(dialogTitle)
        {
            _okButtonUserText = okButtonText ?? ""; // This value will be used with the custom validation button, that does not allow null values.
            _cancelButtonUserText = cancelButtonText; // For this one, setting null would simply reset the label to its default value.
            InvalidSelectionMessage = invalidSelectionMessage;
            if (useSaneSelectionValidation)
            {
                ValidateSelection += ValidateSelectionDefaultHandler;
            }

            // For Open file dialog, allow read only files.
            EnsureReadOnly = true;

            _customValidationButton = new CommonFileDialogButton();
            _customValidationButton.IsProminent = true; // This is the default, but only if there is only one custom control.
            _customValidationButton.Click += _customValidationButton_Click;
            Controls.Add(_customValidationButton);

            // Set constant names for the controls so that we can find them easily:
            // Note: This is done in the ctor. Meaning that showing the same dialog twice won't work.
            // There are other things preventing this scenario anyway. For example,
            // custom controls don't seem to work properly when the dialog is shown for the second time.

            SetFileNameLabel(FileNameLabelTextId);
            SetOkButtonLabel(OkButtonTextId);
            SetCancelButtonLabel(CancelButtonTextId);
            _customValidationButton.Text = CustomValidationButtonTextId;
        }


        /// <summary>
        /// Gets a collection of the selected file names.
        /// </summary>
        public IEnumerable<string> FileNames // Implementation copied from CommonOpenFileDialog
        {
            get
            {
                CheckFileNamesAvailable();
                return base.FileNameCollection;
            }
        }

        /// <summary>
        /// Gets a collection of the selected items as ShellObject objects.
        /// </summary>
        public ICollection<ShellObject> FilesAsShellObject // Implementation copied from CommonOpenFileDialog
        {
            get
            {
                // Check if we have selected files from the user.
                CheckFileItemsAvailable();

                // temp collection to hold our shellobjects
                ICollection<ShellObject> resultItems = new Collection<ShellObject>();

                // Loop through our existing list of filenames, and try to create a concrete type of
                // ShellObject (e.g. ShellLibrary, FileSystemFolder, ShellFile, etc)
                foreach (var si in items)
                {
                    resultItems.Add(ShellObjectFactory.Create(si));
                }

                return resultItems;
            }
        }

        /// <summary>
        /// The handler of our custom validation button, that replaces the normal Ok button.
        /// </summary>
        private void _customValidationButton_Click(object sender, EventArgs e)
        {
            DebugWriteLine("CustomValidation_Click");
            var eventArgs = new CancelEventArgs();
            // The custom selection validation is done in OnFileOk.
            OnFileOk(eventArgs);

            if (!eventArgs.Cancel)
            {
                DebugWriteLine("CustomValidation_Click: Closing the window with HResult.Ok");
                _openDialogCoClass.Close((int)HResult.Ok);
            }
        }

        /// <summary>
        /// This method is called either automatically by the dialog, e.g.: when double clicking on an item,
        /// or from our custom validation button handler.
        /// This ensures the logic flow is as close as possible to the default one.
        /// </summary>
        protected override void OnFileOk(CancelEventArgs e)
        {
            DebugWriteLine("OnFileOk");
            if (OnValidateSelection())
            {
                DebugWriteLine("base.OnFileOk");
                base.OnFileOk(e);
            }
            else
            {
                // If our custom validation fails, we don't even call the base OnFileOk.
                e.Cancel = true;
            }
            DebugWriteLine($"OnFileOk: {!e.Cancel}");
        }

        /// <summary>
        /// This method implements the custom validation logic for files and folders selections.
        /// </summary>
        private bool OnValidateSelection()
        {
            DebugWriteLine("OnValidateSelection");
            // Save the dialog results. This has to be done here because we can't get a reference to IFolderView2 once the dialog has closed.
            var folderView = GetIFolderView2(_openDialogCoClass);
            folderView.GetSelection(true, out var resultArray);

            var eventArgs = new ValidateSelectionEventArgs(new ShellObjectCollection(resultArray, readOnly: true));
            ValidateSelection?.Invoke(this, eventArgs);

            _userFilteredSelection = eventArgs.GetUserFilteredSelection();
            var filteredSelectionHasItems = _userFilteredSelection.Any();
            DebugWriteLine($"OnValidateSelection: Selection valid:{eventArgs.IsSelectionValid}, Filtered selection has items:{filteredSelectionHasItems}");

            if (!eventArgs.IsSelectionValid || !filteredSelectionHasItems)
            {
                var dialogWindowHandle = GetDialogWindowHandle(_openDialogCoClass);
                var owner = new WindowHandleWrapper(dialogWindowHandle);

                MessageBox.Show(owner, InvalidSelectionMessage ?? DefaultInvalidSelectionMessage, null, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            return eventArgs.IsSelectionValid && filteredSelectionHasItems;
        }

        /// <summary>
        /// A default implementation for the <see cref="ValidateSelection"/> event,
        /// that disallows special/invalid folders and files, and dereference .lnk shortcuts.
        /// </summary>
        private static void ValidateSelectionDefaultHandler(object sender, ValidateSelectionEventArgs e)
        {
            /*
             * We want to dereference only the shortcuts, not the other kinds of links.
             * (Uploading a soft link actually uploads the underlying file as expected, while uploading a shortcut doesn't work)
             * Tested link types: 
             * - Shortcut to a file or a folder (.lnk)
             * - Custom shell link (click once shortcut) (.appref-ms)
             * - soft link to a file or folder (.symlink, though only for files, and the extension never shows anywhere)
             * - hard link to a file (this is supposed to be the same thing as an actual file)
             * - directory junction (not sure about what this is)
             * Note: the default open file dialog window dereferences shortcuts and symlinks,
             *       but accessing a symlink returns the file content, unlike accessing a .lnk where the shortcut metadata is returned,
             *       so I don't think that's such a good idea to dereference symlinks...
             * Note 2: the System.Link.TargetParsingPath property is null for a directory junction,
             *         even though IsLink is true, and I didn't find any other property containing the target path
             */

            bool ShouldDereference(ShellObject obj)
            {
                try
                {
                    return obj.IsFileSystemObject && obj.Properties.System.FileExtension.Value?.ToLowerInvariant() == ".lnk";
                }
                catch
                {
                    return false;
                }
            }

            ShellObject GetShellLinkTarget(ShellObject shellLink)
            {
                try
                {
                    // This property seems to work for all types of links (.lnk, appref-ms, symlink...)
                    // See https://blogs.msdn.microsoft.com/oldnewthing/20100702-00/?p=13523 for a potential alternative (native code)
                    var target = shellLink.Properties.GetProperty<string>("System.Link.TargetParsingPath").Value;
                    return target == null ? null : ShellObject.FromParsingName(target);
                }
                catch
                {
                    return null;
                }
            }

            List<ShellObject> GetFilteredSelection(IEnumerable<ShellObject> selection, out bool hasInvalidItems)
            {
                bool IsInvalid(ShellObject obj, ref bool isInvalidAmongItems)
                {
                    // According to the documentation of ShellNativeMethods.FileOpenOptions.AllNonStorageItems,
                    // the dialog default is to allow items with either FileSystem or Stream attributes to be returned.
                    var isInvalid = obj == null || !(obj.IsFileSystemObject || obj.IsStream);
                    if (isInvalid)
                    {
                        // Only override the value if true, never set it to false.
                        isInvalidAmongItems = true;
                    }
                    return isInvalid;
                }

                var result = new List<ShellObject>();
                hasInvalidItems = false;
                foreach (var item in selection)
                {
                    if (!IsInvalid(item, ref hasInvalidItems))
                    {
                        var dereferencedItem = ShouldDereference(item) ? GetShellLinkTarget(item) : item;
                        // If the user selects a shortcut that points to an invalid item, we consider the selection as invalid.
                        if (!IsInvalid(dereferencedItem, ref hasInvalidItems))
                        {
                            result.Add(dereferencedItem);
                        }
                    }
                }
                return result;
            }

            var filteredSelection = GetFilteredSelection(e.Selection, out var selectionHasInvalidItems);
            e.IsSelectionValid = !selectionHasInvalidItems;

            e.OverrideSelection(filteredSelection);
        }

        private static IFolderView2 GetIFolderView2(IFileDialog pfd)
        {
            // Getting an IFolderView2 from the dialog is inspired from this native sample:
            // https://github.com/Microsoft/Windows-classic-samples/blob/master/Samples/Win7Samples/winui/shell/appplatform/CommonFileDialogModes/CommonFileDialogModes.cpp

            var dialogPointer = Marshal.GetIUnknownForObject(pfd);
            IFolderView2 resultFolderView2;
            try
            {
                var folderViewGuid = new Guid(ExplorerBrowserIIDGuid.IFolderView);
                var folderView2Guid = new Guid(ExplorerBrowserIIDGuid.IFolderView2);

                // First magic bit:
                var queryServiceResult = ShellNativeMethods.IUnknown_QueryService(dialogPointer, ref folderViewGuid, ref folderView2Guid, out var requestedInterfacePointer);
                if (queryServiceResult != HResult.Ok)
                {
                    throw new Win32Exception((int)queryServiceResult);
                }

                resultFolderView2 = (IFolderView2)Marshal.GetObjectForIUnknown(requestedInterfacePointer);
                Marshal.Release(requestedInterfacePointer);
            }
            finally
            {
                Marshal.Release(dialogPointer);
            }

            return resultFolderView2;
        }

        private static IntPtr GetDialogWindowHandle(IFileDialog pfd)
        {
            // Second magic bit:
            var oleWindow = (IOleWindow)pfd;
            oleWindow.GetWindow(out var dialogWindowHandle);
            return dialogWindowHandle;
        }

        /// <inheritdoc />
        protected override void OnOpening(EventArgs e)
        {
            base.OnOpening(e);
            InitializeHook();
        }

        private void InitializeHook()
        {
            var dialogWindowHandle = GetDialogWindowHandle(_openDialogCoClass);

            // We have the window handle. Now we need the child controls:

            var okButton = WindowNativeMethods.GetAllChildrenWindowHandles(dialogWindowHandle, className: "Button", windowTitle: OkButtonTextId).FirstOrDefault();
            var cancelButton = WindowNativeMethods.GetAllChildrenWindowHandles(dialogWindowHandle, className: "Button", windowTitle: CancelButtonTextId).FirstOrDefault();
            var customValidationButton = WindowNativeMethods.GetAllChildrenWindowHandles(dialogWindowHandle, className: "Button", windowTitle: CustomValidationButtonTextId).FirstOrDefault();
            var fileNameCombBox = WindowNativeMethods.GetAllChildrenWindowHandles(dialogWindowHandle, className: "ComboBoxEx32").FirstOrDefault();
            var fileNameLabel = WindowNativeMethods.GetAllChildrenWindowHandles(dialogWindowHandle, className: "Static", windowTitle: FileNameLabelTextId).FirstOrDefault();
            var explorerView = WindowNativeMethods.GetAllChildrenWindowHandles(dialogWindowHandle, className: "DUIViewWndClassName").FirstOrDefault();

            const string exceptionMessage = "Could not find the dialog child control ";
            if (okButton == default(IntPtr)) throw new Exception(exceptionMessage + nameof(okButton));
            if (cancelButton == default(IntPtr)) throw new Exception(exceptionMessage + nameof(cancelButton));
            if (customValidationButton == default(IntPtr)) throw new Exception(exceptionMessage + nameof(customValidationButton));
            if (fileNameCombBox == default(IntPtr)) throw new Exception(exceptionMessage + nameof(fileNameCombBox));
            if (fileNameLabel == default(IntPtr)) throw new Exception(exceptionMessage + nameof(fileNameLabel));
            if (explorerView == default(IntPtr)) throw new Exception(exceptionMessage + nameof(explorerView));

            // Reset the visible labels to their desired values, now that we have the control handles:

            SetCancelButtonLabel(_cancelButtonUserText);
            _customValidationButton.Text = _okButtonUserText;

            // Measurements:

            NativeRect okButtonSize;
            int comboBoxMaxHeight;
            int customValidationButtonDesiredX;
            int customValidationButtonDesiredY;

            // This method will be called each time the window is resized
            void RefreshMeasurements()
            {
                WindowNativeMethods.GetWindowRect(okButton, out okButtonSize);
                WindowNativeMethods.GetWindowRect(cancelButton, out var cancelButtonSize);
                WindowNativeMethods.GetWindowRect(customValidationButton, out var customValidationButtonSize);
                WindowNativeMethods.GetWindowRect(fileNameCombBox, out var fileNameCombBoxSize);
                WindowNativeMethods.GetWindowRect(fileNameLabel, out var fileNameLabelSize);

                // https://stackoverflow.com/questions/18034975/how-do-i-find-position-of-a-win32-control-window-relative-to-its-parent-window
                const int pointsCount = 2;// 1 is for points, 2 for rects
                WindowNativeMethods.MapWindowPoints(IntPtr.Zero, dialogWindowHandle, ref okButtonSize, pointsCount);
                WindowNativeMethods.MapWindowPoints(IntPtr.Zero, dialogWindowHandle, ref cancelButtonSize, pointsCount);
                WindowNativeMethods.MapWindowPoints(IntPtr.Zero, dialogWindowHandle, ref customValidationButtonSize, pointsCount);
                WindowNativeMethods.MapWindowPoints(IntPtr.Zero, dialogWindowHandle, ref fileNameCombBoxSize, pointsCount);
                WindowNativeMethods.MapWindowPoints(IntPtr.Zero, dialogWindowHandle, ref fileNameLabelSize, pointsCount);

                var fileNameCombBoxHeight = fileNameCombBoxSize.Bottom - fileNameCombBoxSize.Top;
                var fileNameLabelHeight = fileNameLabelSize.Bottom - fileNameLabelSize.Top;
                comboBoxMaxHeight = Math.Max(fileNameCombBoxHeight, fileNameLabelHeight);

                var okCancelButtonsMargin = cancelButtonSize.Left - okButtonSize.Right;
                var customValidationButtonWidth = customValidationButtonSize.Right - customValidationButtonSize.Left;

                customValidationButtonDesiredX = cancelButtonSize.Left - okCancelButtonsMargin - customValidationButtonWidth;
                customValidationButtonDesiredY = cancelButtonSize.Top;
            }

            RefreshMeasurements();

            // Hide controls we don't want to use:

            WindowNativeMethods.SetWindowPos(fileNameCombBox, IntPtr.Zero, 0, 0, 0, 0,
                SetWindowPosFlags.SWP_NOZORDER | SetWindowPosFlags.SWP_NOSIZE | SetWindowPosFlags.SWP_HIDEWINDOW);
            WindowNativeMethods.SetWindowPos(fileNameLabel, IntPtr.Zero, 0, 0, 0, 0,
                 SetWindowPosFlags.SWP_NOZORDER | SetWindowPosFlags.SWP_NOSIZE | SetWindowPosFlags.SWP_HIDEWINDOW);

            // Not sure why, but moving the ok button up increases the explorer view height.
            // The explorer view then follows the window when it resizes, as desired.
            WindowNativeMethods.SetWindowPos(
                okButton, IntPtr.Zero, okButtonSize.Left, okButtonSize.Top - comboBoxMaxHeight, 0, 0,
                SetWindowPosFlags.SWP_NOZORDER | SetWindowPosFlags.SWP_NOSIZE | SetWindowPosFlags.SWP_HIDEWINDOW);

            // Now the only thing we need to to regarding the dialog layout,
            // is moving our custom validation button closer to the cancel button, where the ok button is supposed to be.
            // Unfortunately, to do it properly, we have to hook into the dialog message loop
            // so that we can readjust the button position if the form is resized, and its layout recomputed.

            var hookPtr = IntPtr.Zero;
            _callback = CallWndRetProc;
            hookPtr = WndProcRetHookNativeMethods.SetWindowsHookEx(_callback);

            IntPtr CallWndRetProc(int code, IntPtr wParam, CWPRETSTRUCT obj)
            {
                switch ((WindowMessage)obj.message)
                {
                    case WindowMessage.Destroy:
                        if (obj.hWnd == dialogWindowHandle)
                        {
                            WndProcRetHookNativeMethods.UnhookWindowsHookEx(hookPtr);
                        }
                        break;
                    case WindowMessage.ShowWindow:
                    case WindowMessage.Size:
                        RefreshMeasurements();
                        goto case WindowMessage.Paint;
                    case WindowMessage.Paint:
                    case WindowMessage.NCPaint:
                        // Doing this for each repaint is probably overkill, but I'm not sure how to it more efficiently...
                        WindowNativeMethods.SetWindowPos(
                            customValidationButton, IntPtr.Zero, customValidationButtonDesiredX, customValidationButtonDesiredY, 0, 0,
                            SetWindowPosFlags.SWP_NOZORDER | SetWindowPosFlags.SWP_NOSIZE);
                        break;
                }
                return WndProcRetHookNativeMethods.CallNextHookEx(IntPtr.Zero, code, wParam, obj);
            }
        }

        internal override void InitializeNativeFileDialog()
        {
            if (_openDialogCoClass == null)
            {
                _openDialogCoClass = new NativeFileOpenDialog();
            }
        }

        internal override IFileDialog GetNativeFileDialog()
        {
            Debug.Assert(_openDialogCoClass != null, "Must call Initialize() before fetching dialog interface");
            return _openDialogCoClass;
        }

        internal override void CleanUpNativeFileDialog()
        {
            if (_openDialogCoClass != null)
            {
                Marshal.ReleaseComObject(_openDialogCoClass);
                _openDialogCoClass = null;
            }
            _userFilteredSelection = null;
        }

        internal override void PopulateWithFileNames(Collection<string> names) => PopulateItems(names, null);
        internal override void PopulateWithIShellItems(Collection<IShellItem> shellItems) => PopulateItems(null, shellItems);
        private void PopulateItems(ICollection<string> names, ICollection<IShellItem> shellItems)
        {
            names?.Clear();
            shellItems?.Clear();

            if (_userFilteredSelection == null)
            {
                return;
            }

            foreach (var shellItem in _userFilteredSelection)
            {
                shellItems?.Add(shellItem.NativeShellItem);
                names?.Add(GetFileNameFromShellItem(shellItem.NativeShellItem));
            }
        }

        internal override ShellNativeMethods.FileOpenOptions GetDerivedOptionFlags(ShellNativeMethods.FileOpenOptions flags)
        {
            flags |= ShellNativeMethods.FileOpenOptions.AllowMultiSelect
                     | ShellNativeMethods.FileOpenOptions.NoValidate;

            return flags;
        }

        [Conditional("DEBUG_OFOFD")]
        private static void DebugWriteLine(object message) => Trace.WriteLine(message);

        private class WindowHandleWrapper : IWin32Window
        {
            public IntPtr Handle { get; }
            public WindowHandleWrapper(IntPtr handle) => Handle = handle;
        }
    }

    /// <summary>
    /// Event argument for the ValidateSelection event.
    /// </summary>
    public class ValidateSelectionEventArgs : EventArgs
    {
        /// <summary>
        /// The original collection of <see cref="ShellObject"/> items selected by the user.
        /// </summary>
        public IEnumerable<ShellObject> Selection { get; }

        /// <summary>
        /// If set to false, the dialog will stay opened and a message will be displayed
        /// to indicate the user must change their selection.
        /// </summary>
        public bool IsSelectionValid { get; set; } = true;

        private List<ShellObject> _userFilteredSelection;

        /// <summary>
        /// Creates a new instance of this class.
        /// </summary>
        /// <param name="selection"></param>
        public ValidateSelectionEventArgs(IEnumerable<ShellObject> selection)
            => Selection = selection;

        /// <summary>
        /// The provided collection will be as the return value of the dialog instead of the original selection.
        /// </summary>
        /// <param name="filteredSelection"></param>
        public void OverrideSelection(IEnumerable<ShellObject> filteredSelection)
            => _userFilteredSelection = filteredSelection.ToList();
        internal ICollection<ShellObject> GetUserFilteredSelection()
            => _userFilteredSelection ?? (_userFilteredSelection = Selection.ToList());
    }
}

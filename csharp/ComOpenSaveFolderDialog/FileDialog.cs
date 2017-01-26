/*  Company :       Nequeo Pty Ltd, http://www.nequeo.com.au/
 *  Copyright :     Copyright © Nequeo Pty Ltd 2015 http://www.nequeo.com.au/
 * 
 *  File :          
 *  Purpose :       
 * 
 */

#region Nequeo Pty Ltd License
/*
    Permission is hereby granted, free of charge, to any person
    obtaining a copy of this software and associated documentation
    files (the "Software"), to deal in the Software without
    restriction, including without limitation the rights to use,
    copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the
    Software is furnished to do so, subject to the following
    conditions:

    The above copyright notice and this permission notice shall be
    included in all copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
    EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
    OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
    NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
    HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
    WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
    FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
    OTHER DEALINGS IN THE SOFTWARE.
*/
#endregion

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Nequeo.Interop
{
    /// <summary>
    /// Open and Save file dialog, using COM import of native IFileOpenDialog and IFileSaveDialog interfaces.
    /// </summary>
    public class FileDialog
    {
        /// <summary>
        /// Show the open file dialog.
        /// </summary>
        /// <param name="title">The title.</param>
        /// <param name="okButtonText">The OK button text.</param>
        /// <param name="fileNameText">The selected file name text.</param>
        /// <param name="fileName">The default file name.</param>
        /// <param name="defaultFolder">The default folder.</param>
        /// <param name="defaultExtension">The default extension.</param>
        /// <param name="allowMultiSelect">True to allow multi-select files.</param>
        /// <param name="showHiddenFiles">True to display hidden files.</param>
        /// <param name="pathMustExist">True if path must exist.</param>
        /// <param name="fileMustExist">True if file must exist.</param>
        /// <param name="noTestFileCreated">True if no test is done for file created</param>
        /// <param name="displayOverWritePrompt">True to display over write file prompt.</param>
        /// <param name="filters">The collection of filters.</param>
        /// <returns>The collection of file URLs.</returns>
        public static string[] ShowOpenFileDialog(string title = "Open", string okButtonText = "Open", string fileNameText = "File name:", string fileName = null,
            string defaultFolder = null, string defaultExtension = null, bool allowMultiSelect = false, bool showHiddenFiles = false, bool pathMustExist = false,
            bool fileMustExist = false, bool noTestFileCreated = false, bool displayOverWritePrompt = false, FileDialogFilter[] filters = null)
        {
            bool optionSet = false;
            bool removeOverWrite = false;
            List<string> files = new List<string>();
            FileDialogNative.IFileOpenDialog dialog = null;

            try
            {
                // Create the COM object.
                dialog = new FileDialogNative.NativeFileOpenDialog();
                dialog.SetTitle(title);
                dialog.SetOkButtonLabel(okButtonText);
                dialog.SetFileNameLabel(fileNameText);

                // Set file name initially.
                if (!string.IsNullOrEmpty(fileName))
                    dialog.SetFileName(fileName);

                // Set default extension initially.
                if (!string.IsNullOrEmpty(defaultExtension))
                    dialog.SetFileName(defaultExtension);

                // Set overwrite prompt.
                FileDialogNative.FOS fos = FileDialogNative.FOS.FOS_OVERWRITEPROMPT;

                // Allow over write prompt select.
                if (displayOverWritePrompt)
                {
                    fos |= FileDialogNative.FOS.FOS_OVERWRITEPROMPT;
                    optionSet = true;
                }
                else
                {
                    // Remove this option.
                    removeOverWrite = true;
                }

                // Allow no test file created select.
                if (noTestFileCreated)
                {
                    fos |= FileDialogNative.FOS.FOS_NOTESTFILECREATE;
                    optionSet = true;
                }

                // Allow path must exist select.
                if (pathMustExist)
                {
                    fos |= FileDialogNative.FOS.FOS_PATHMUSTEXIST;
                    optionSet = true;
                }

                // Allow file must exist select.
                if (fileMustExist)
                {
                    fos |= FileDialogNative.FOS.FOS_FILEMUSTEXIST;
                    optionSet = true;
                }

                // Allow multi select.
                if (allowMultiSelect)
                {
                    fos |= FileDialogNative.FOS.FOS_ALLOWMULTISELECT;
                    optionSet = true;
                }

                // Show hidden folders.
                if (showHiddenFiles)
                {
                    fos |= FileDialogNative.FOS.FOS_FORCESHOWHIDDEN;
                    optionSet = true;
                }

                // Set the options
                if (optionSet)
                {
                    // Remove over write first.
                    if (removeOverWrite)
                        fos &= ~FileDialogNative.FOS.FOS_OVERWRITEPROMPT;

                    dialog.SetOptions(fos);
                }

                // Set the default folder.
                if (!String.IsNullOrEmpty(defaultFolder))
                {
                    //TODO
                }

                // Set filters.
                if (filters != null && filters.Length > 0)
                {
                    // Collection of filters.
                    List<FileDialogNative.COMDLG_FILTERSPEC> listFilters = new List<FileDialogNative.COMDLG_FILTERSPEC>();

                    // For each filter.
                    for (int i = 0; i < filters.Length; i++)
                    {
                        // Get the current filter.
                        FileDialogFilter filter = filters[i];

                        // Validate data.
                        if (!string.IsNullOrEmpty(filter.FilterName) && !string.IsNullOrEmpty(filter.FilterValue))
                        {
                            // Add.
                            listFilters.Add(new FileDialogNative.COMDLG_FILTERSPEC() { pszName = filter.FilterName, pszSpec = filter.FilterValue });
                        }
                    }

                    // If filters exist.
                    if (listFilters.Count > 0)
                        dialog.SetFileTypes((uint)listFilters.Count, listFilters.ToArray());
                }

                // Show the dialog.
                int ret = dialog.Show(IntPtr.Zero);

                // If OK.
                if (ret == 0)
                {
                    // Get items.
                    FileDialogNative.IShellItemArray items = null;
                    dialog.GetSelectedItems(out items);

                    // Not null.
                    if (items != null)
                    {
                        // Get the number of items selected.
                        uint count = 0;
                        items.GetCount(out count);

                        // If items selected.
                        if (count > 0)
                        {
                            // For each item selected.
                            for (uint i = 0; i < count; i++)
                            {
                                // Get item
                                FileDialogNative.IShellItem item = null;
                                items.GetItemAt(i, out item);

                                // Get URL.
                                IntPtr ptrFileURL = IntPtr.Zero;
                                item.GetDisplayName(FileDialogNative.SIGDN.SIGDN_URL, out ptrFileURL);
                                string fileURL = Marshal.PtrToStringAuto(ptrFileURL);

                                // Add files to the collection.
                                files.Add(fileURL);
                            }
                        }
                    }
                }
            }
            finally
            {
                // Free the com object.
                if (dialog != null)
                    Marshal.FinalReleaseComObject(dialog);
            }

            // Return the list of selected files.
            return files.ToArray();
        }

        /// <summary>
        /// Show the save file dialog.
        /// </summary>
        /// <param name="title">The title.</param>
        /// <param name="okButtonText">The OK button text.</param>
        /// <param name="fileNameText">The selected file name text.</param>
        /// <param name="fileName">The default file name.</param>
        /// <param name="defaultFolder">The default folder.</param>
        /// <param name="defaultExtension">The default extension.</param>
        /// <param name="allowMultiSelect">True to allow multi-select files.</param>
        /// <param name="showHiddenFiles">True to display hidden files.</param>
        /// <param name="pathMustExist">True if path must exist.</param>
        /// <param name="fileMustExist">True if file must exist.</param>
        /// <param name="noTestFileCreated">True if no test is done for file created</param>
        /// <param name="displayOverWritePrompt">True to display over write file prompt.</param>
        /// <param name="filters">The collection of filters.</param>
        /// <returns>The selected URL.</returns>
        public static string ShowSaveFileDialog(string title = "Save As", string okButtonText = "Save", string fileNameText = "File name:", string fileName = null,
            string defaultFolder = null, string defaultExtension = null, bool allowMultiSelect = false, bool showHiddenFiles = false, bool pathMustExist = false,
            bool fileMustExist = false, bool noTestFileCreated = false, bool displayOverWritePrompt = false, FileDialogFilter[] filters = null)
        {
            bool optionSet = false;
            bool removeOverWrite = false;
            string file = null;
            FileDialogNative.IFileSaveDialog dialog = null;

            try
            {
                // Create the COM object.
                dialog = new FileDialogNative.NativeFileSaveDialog();
                dialog.SetTitle(title);
                dialog.SetOkButtonLabel(okButtonText);
                dialog.SetFileNameLabel(fileNameText);

                // Set file name initially.
                if (!string.IsNullOrEmpty(fileName))
                    dialog.SetFileName(fileName);

                // Set default extension initially.
                if (!string.IsNullOrEmpty(defaultExtension))
                    dialog.SetFileName(defaultExtension);

                // Set overwrite prompt.
                FileDialogNative.FOS fos = FileDialogNative.FOS.FOS_OVERWRITEPROMPT;

                // Allow over write prompt select.
                if (displayOverWritePrompt)
                {
                    fos |= FileDialogNative.FOS.FOS_OVERWRITEPROMPT;
                    optionSet = true;
                }
                else
                {
                    // Remove this option.
                    removeOverWrite = true;
                }

                // Allow no test file created select.
                if (noTestFileCreated)
                {
                    fos |= FileDialogNative.FOS.FOS_NOTESTFILECREATE;
                    optionSet = true;
                }

                // Allow path must exist select.
                if (pathMustExist)
                {
                    fos |= FileDialogNative.FOS.FOS_PATHMUSTEXIST;
                    optionSet = true;
                }

                // Allow file must exist select.
                if (fileMustExist)
                {
                    fos |= FileDialogNative.FOS.FOS_FILEMUSTEXIST;
                    optionSet = true;
                }

                // Allow multi select.
                if (allowMultiSelect)
                {
                    fos |= FileDialogNative.FOS.FOS_ALLOWMULTISELECT;
                    optionSet = true;
                }

                // Show hidden folders.
                if (showHiddenFiles)
                {
                    fos |= FileDialogNative.FOS.FOS_FORCESHOWHIDDEN;
                    optionSet = true;
                }

                // Set the options
                if (optionSet)
                {
                    // Remove over write first.
                    if (removeOverWrite)
                        fos &= ~FileDialogNative.FOS.FOS_OVERWRITEPROMPT;

                    dialog.SetOptions(fos);
                }

                // Set the default folder.
                if (!String.IsNullOrEmpty(defaultFolder))
                {
                    //TODO
                }

                // Set filters.
                if (filters != null && filters.Length > 0)
                {
                    // Collection of filters.
                    List<FileDialogNative.COMDLG_FILTERSPEC> listFilters = new List<FileDialogNative.COMDLG_FILTERSPEC>();

                    // For each filter.
                    for (int i = 0; i < filters.Length; i++)
                    {
                        // Get the current filter.
                        FileDialogFilter filter = filters[i];

                        // Validate data.
                        if (!string.IsNullOrEmpty(filter.FilterName) && !string.IsNullOrEmpty(filter.FilterValue))
                        {
                            // Add.
                            listFilters.Add(new FileDialogNative.COMDLG_FILTERSPEC() { pszName = filter.FilterName, pszSpec = filter.FilterValue });
                        }
                    }

                    // If filters exist.
                    if (listFilters.Count > 0)
                        dialog.SetFileTypes((uint)listFilters.Count, listFilters.ToArray());
                }

                // Show the dialog.
                int ret = dialog.Show(IntPtr.Zero);

                // If OK.
                if (ret == 0)
                {
                    // Get item.
                    FileDialogNative.IShellItem item = null;
                    dialog.GetResult(out item);

                    // Not null.
                    if (item != null)
                    {
                        // Get URL.
                        IntPtr ptrFileURL = IntPtr.Zero;
                        item.GetDisplayName(FileDialogNative.SIGDN.SIGDN_URL, out ptrFileURL);
                        string fileURL = Marshal.PtrToStringAuto(ptrFileURL);

                        // Assign the URL.
                        file = fileURL;
                    }
                }
            }
            finally
            {
                // Free the com object.
                if (dialog != null)
                    Marshal.FinalReleaseComObject(dialog);
            }

            // Return the selected file.
            return file;
        }
    }

    /// <summary>
    /// File dialog filter.
    /// </summary>
    public class FileDialogFilter
    {
        /// <summary>
        /// Gets or sets the filter name (e.g. Image Files)
        /// </summary>
        public string FilterName { get; set; }

        /// <summary>
        /// Gets or sets the filter value (e.g. *.bmp;*.jpg)
        /// </summary>
        public string FilterValue { get; set; }
    }
}

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
    /// Folder dialog, using COM import of native IShellFolder interfaces. 
    /// </summary>
    public class FolderDialog
    {
        /// <summary>
        /// Show the folder or file.
        /// </summary>
        /// <param name="path">The path to the file or folder.</param>
        /// <param name="edit">True to edit the file; else false.</param>
        public static void ShowFileOrFolder(string path, bool edit = false)
        {
            FolderDialogNative.FileOrFolder(path, edit);
        }

        /// <summary>
        /// Show the folders or files.
        /// </summary>
        /// <param name="path">The path to the file or folder.</param>
        /// <param name="filenames">The list of files to select.</param>
        public static void ShowFilesOrFolders(string path, ICollection<string> filenames)
        {
            FolderDialogNative.FilesOrFolders(path, filenames);
        }

        /// <summary>
        /// Show the folders or files.
        /// </summary>
        /// <param name="paths">The collection of folder and files to show.</param>
        public static void ShowFilesOrFolders(params string[] paths)
        {
            FolderDialogNative.FilesOrFolders(paths);
        }

        /// <summary>
        /// Show the folder browser dialog.
        /// </summary>
        /// <param name="caption">The dialog caption.</param>
        /// <param name="initialPath">The initial path to navigate to.</param>
        public static void ShowFolderDialog(string caption, string initialPath)
        {
            FolderDialogNative.BrowseForFolder folder = new FolderDialogNative.BrowseForFolder();
            IntPtr parent = IntPtr.Zero;
            folder.SelectFolder(caption, initialPath, parent);
        }

        /// <summary>
        /// Show open folder dialog.
        /// </summary>
        /// <param name="title">The dialog title.</param>
        /// <param name="okButtonText">The OK button text.</param>
        /// <param name="folderNameText">The folder name text.</param>
        /// <param name="defaultFolder">The default path.</param>
        /// <param name="allowMultiSelect">True to allow multi-select folders.</param>
        /// <param name="showHiddenFolders">True to show hidden foldes.</param>
        /// <returns>The collection of folder URLs.</returns>
        public static string[] ShowOpenFolderDialog(string title = "Open", string okButtonText = "Open", string folderNameText = "Folder:",
            string defaultFolder = null, bool allowMultiSelect = false, bool showHiddenFolders = false)
        {
            List<string> folders = new List<string>();
            FileDialogNative.IFileOpenDialog dialog = null;

            try
            {
                // Create the COM object.
                dialog = new FileDialogNative.NativeFileOpenDialog();
                dialog.SetTitle(title);
                dialog.SetOkButtonLabel(okButtonText);
                dialog.SetFileNameLabel(folderNameText);

                // Set folder picker.
                FileDialogNative.FOS fos = FileDialogNative.FOS.FOS_PICKFOLDERS;

                // Allow multi select.
                if (allowMultiSelect)
                    fos |= FileDialogNative.FOS.FOS_ALLOWMULTISELECT;

                // Show hidden folders.
                if (showHiddenFolders)
                    fos |= FileDialogNative.FOS.FOS_FORCESHOWHIDDEN;

                // Set the options
                dialog.SetOptions(fos);

                // Set the default folder.
                if (!String.IsNullOrEmpty(defaultFolder))
                {
                    //TODO
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
                                IntPtr ptrFolderURL = IntPtr.Zero;
                                item.GetDisplayName(FileDialogNative.SIGDN.SIGDN_URL, out ptrFolderURL);
                                string folderURL = Marshal.PtrToStringAuto(ptrFolderURL);

                                // Add folder to the collection.
                                folders.Add(folderURL);
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

            // Return the list of selected folders.
            return folders.ToArray();
        }
    }
}

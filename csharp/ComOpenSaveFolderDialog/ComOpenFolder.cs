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
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace Nequeo.Interop
{
    /// <summary>
    /// Open folder COM dialog interface.
    /// </summary>
    static class FolderDialogNative
    {
        /// <summary>
        /// 
        /// </summary>
        [Flags]
        enum SHCONT : ushort
        {
            SHCONTF_CHECKING_FOR_CHILDREN = 0x0010,
            SHCONTF_FOLDERS = 0x0020,
            SHCONTF_NONFOLDERS = 0x0040,
            SHCONTF_INCLUDEHIDDEN = 0x0080,
            SHCONTF_INIT_ON_FIRST_NEXT = 0x0100,
            SHCONTF_NETPRINTERSRCH = 0x0200,
            SHCONTF_SHAREABLE = 0x0400,
            SHCONTF_STORAGE = 0x0800,
            SHCONTF_NAVIGATION_ENUM = 0x1000,
            SHCONTF_FASTITEMS = 0x2000,
            SHCONTF_FLATLIST = 0x4000,
            SHCONTF_ENABLE_ASYNC = 0x8000
        }

        /// <summary>
        /// 
        /// </summary>
        [ComImport,
        Guid("000214E6-0000-0000-C000-000000000046"),
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
        ComConversionLoss]
        interface IShellFolder
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void ParseDisplayName(IntPtr hwnd, [In, MarshalAs(UnmanagedType.Interface)] IBindCtx pbc, [In, MarshalAs(UnmanagedType.LPWStr)] string pszDisplayName, [Out] out uint pchEaten, [Out] out IntPtr ppidl, [In, Out] ref uint pdwAttributes);
            [PreserveSig]
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int EnumObjects([In] IntPtr hwnd, [In] SHCONT grfFlags, [MarshalAs(UnmanagedType.Interface)] out IEnumIDList ppenumIDList);

            [PreserveSig]
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int BindToObject([In] IntPtr pidl, [In, MarshalAs(UnmanagedType.Interface)] IBindCtx pbc, [In] ref Guid riid, [Out, MarshalAs(UnmanagedType.Interface)] out IShellFolder ppv);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void BindToStorage([In] ref IntPtr pidl, [In, MarshalAs(UnmanagedType.Interface)] IBindCtx pbc, [In] ref Guid riid, out IntPtr ppv);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void CompareIDs([In] IntPtr lParam, [In] ref IntPtr pidl1, [In] ref IntPtr pidl2);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void CreateViewObject([In] IntPtr hwndOwner, [In] ref Guid riid, out IntPtr ppv);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void GetAttributesOf([In] uint cidl, [In] IntPtr apidl, [In, Out] ref uint rgfInOut);


            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void GetUIObjectOf([In] IntPtr hwndOwner, [In] uint cidl, [In] IntPtr apidl, [In] ref Guid riid, [In, Out] ref uint rgfReserved, out IntPtr ppv);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void GetDisplayNameOf([In] ref IntPtr pidl, [In] uint uFlags, out IntPtr pName);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void SetNameOf([In] IntPtr hwnd, [In] ref IntPtr pidl, [In, MarshalAs(UnmanagedType.LPWStr)] string pszName, [In] uint uFlags, [Out] IntPtr ppidlOut);
        }

        /// <summary>
        /// 
        /// </summary>
        [ComImport,
        Guid("000214F2-0000-0000-C000-000000000046"),
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IEnumIDList
        {
            [PreserveSig]
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int Next(uint celt, IntPtr rgelt, out uint pceltFetched);

            [PreserveSig]
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int Skip([In] uint celt);

            [PreserveSig]
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int Reset();

            [PreserveSig]
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int Clone([MarshalAs(UnmanagedType.Interface)] out IEnumIDList ppenum);
        }

        /// <summary>
        /// 
        /// </summary>
        static class NativeMethods
        {
            [DllImport("shell32.dll", EntryPoint = "SHGetDesktopFolder", CharSet = CharSet.Unicode,
                SetLastError = true)]
            static extern int SHGetDesktopFolder_([MarshalAs(UnmanagedType.Interface)] out IShellFolder ppshf);

            public static IShellFolder SHGetDesktopFolder()
            {
                IShellFolder result;
                Marshal.ThrowExceptionForHR(SHGetDesktopFolder_(out result));
                return result;
            }

            [DllImport("shell32.dll", EntryPoint = "SHOpenFolderAndSelectItems")]
            static extern int SHOpenFolderAndSelectItems_(
                [In] IntPtr pidlFolder, uint cidl, [In, Optional, MarshalAs(UnmanagedType.LPArray)] IntPtr[] apidl,
                int dwFlags);

            public static void SHOpenFolderAndSelectItems(IntPtr pidlFolder, IntPtr[] apidl, int dwFlags)
            {
                var cidl = (apidl != null) ? (uint)apidl.Length : 0U;
                var result = SHOpenFolderAndSelectItems_(pidlFolder, cidl, apidl, dwFlags);
                Marshal.ThrowExceptionForHR(result);
            }

            [DllImport("shell32.dll")]
            public static extern void ILFree([In] IntPtr pidl);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="parentFolder"></param>
        /// <param name="displayName"></param>
        /// <returns></returns>
        static IntPtr GetShellFolderChildrenRelativePIDL(IShellFolder parentFolder, string displayName)
        {
            uint pchEaten;
            uint pdwAttributes = 0;
            IntPtr ppidl;
            parentFolder.ParseDisplayName(IntPtr.Zero, null, displayName, out pchEaten, out ppidl, ref pdwAttributes);

            return ppidl;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        static IntPtr PathToAbsolutePIDL(string path)
        {
            var desktopFolder = NativeMethods.SHGetDesktopFolder();
            return GetShellFolderChildrenRelativePIDL(desktopFolder, path);
        }

        static Guid IID_IShellFolder = typeof(IShellFolder).GUID;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="pidl"></param>
        /// <returns></returns>
        static IShellFolder PIDLToShellFolder(IShellFolder parent, IntPtr pidl)
        {
            IShellFolder folder;
            var result = parent.BindToObject(pidl, null, ref IID_IShellFolder, out folder);
            Marshal.ThrowExceptionForHR((int)result);
            return folder;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pidl"></param>
        /// <returns></returns>
        static IShellFolder PIDLToShellFolder(IntPtr pidl)
        {
            return PIDLToShellFolder(NativeMethods.SHGetDesktopFolder(), pidl);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pidlFolder"></param>
        /// <param name="apidl"></param>
        /// <param name="edit"></param>
        static void SHOpenFolderAndSelectItems(IntPtr pidlFolder, IntPtr[] apidl, bool edit)
        {
            NativeMethods.SHOpenFolderAndSelectItems(pidlFolder, apidl, edit ? 1 : 0);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        /// <param name="edit"></param>
        public static void FileOrFolder(string path, bool edit = false)
        {
            if (path == null) throw new ArgumentNullException("path");

            var pidl = PathToAbsolutePIDL(path);
            try
            {
                SHOpenFolderAndSelectItems(pidl, null, edit);
            }
            finally
            {
                NativeMethods.ILFree(pidl);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="paths"></param>
        /// <returns></returns>
        static IEnumerable<FileSystemInfo> PathToFileSystemInfo(IEnumerable<string> paths)
        {
            foreach (var path in paths)
            {
                var fixedPath = path;
                if (fixedPath.EndsWith(Path.DirectorySeparatorChar.ToString())
                    || fixedPath.EndsWith(Path.AltDirectorySeparatorChar.ToString()))
                {
                    fixedPath = fixedPath.Remove(fixedPath.Length - 1);
                }

                if (Directory.Exists(fixedPath))
                {
                    yield return new DirectoryInfo(fixedPath);
                }
                else if (File.Exists(fixedPath))
                {
                    yield return new FileInfo(fixedPath);
                }
                else
                {
                    throw new FileNotFoundException
                        (string.Format("The specified file or folder doesn't exists : {0}", fixedPath),
                        fixedPath);
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="parentDirectory"></param>
        /// <param name="filenames"></param>
        public static void FilesOrFolders(string parentDirectory, ICollection<string> filenames)
        {
            if (filenames == null) throw new ArgumentNullException("filenames");
            if (filenames.Count == 0) return;

            var parentPidl = PathToAbsolutePIDL(parentDirectory);
            try
            {
                var parent = PIDLToShellFolder(parentPidl);
                var filesPidl = filenames
                    .Select(filename => GetShellFolderChildrenRelativePIDL(parent, filename))
                    .ToArray();

                try
                {
                    SHOpenFolderAndSelectItems(parentPidl, filesPidl, false);
                }
                finally
                {
                    foreach (var pidl in filesPidl)
                    {
                        NativeMethods.ILFree(pidl);
                    }
                }
            }
            finally
            {
                NativeMethods.ILFree(parentPidl);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="paths"></param>
        public static void FilesOrFolders(params string[] paths)
        {
            FilesOrFolders((IEnumerable<string>)paths);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="paths"></param>
        public static void FilesOrFolders(IEnumerable<string> paths)
        {
            if (paths == null) throw new ArgumentNullException("paths");

            FilesOrFolders(PathToFileSystemInfo(paths));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="paths"></param>
        public static void FilesOrFolders(IEnumerable<FileSystemInfo> paths)
        {
            if (paths == null) throw new ArgumentNullException("paths");
            var pathsArray = paths.ToArray();
            if (pathsArray.Count() == 0) return;

            var explorerWindows = pathsArray.GroupBy(p => Path.GetDirectoryName(p.FullName));

            foreach (var explorerWindowPaths in explorerWindows)
            {
                var parentDirectory = Path.GetDirectoryName(explorerWindowPaths.First().FullName);
                FilesOrFolders(parentDirectory, explorerWindowPaths.Select(fsi => fsi.Name).ToList());
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public class BrowseForFolder
        {
            // Constants for sending and receiving messages in BrowseCallBackProc
            public const int WM_USER = 0x400;
            public const int BFFM_INITIALIZED = 1;
            public const int BFFM_SELCHANGED = 2;
            public const int BFFM_VALIDATEFAILEDA = 3;
            public const int BFFM_VALIDATEFAILEDW = 4;
            public const int BFFM_IUNKNOWN = 5; // provides IUnknown to client. lParam: IUnknown*
            public const int BFFM_SETSTATUSTEXTA = WM_USER + 100;
            public const int BFFM_ENABLEOK = WM_USER + 101;
            public const int BFFM_SETSELECTIONA = WM_USER + 102;
            public const int BFFM_SETSELECTIONW = WM_USER + 103;
            public const int BFFM_SETSTATUSTEXTW = WM_USER + 104;
            public const int BFFM_SETOKTEXT = WM_USER + 105; // Unicode only
            public const int BFFM_SETEXPANDED = WM_USER + 106; // Unicode only

            // Browsing for directory.
            private uint BIF_RETURNONLYFSDIRS = 0x0001;  // For finding a folder to start document searching
            private uint BIF_DONTGOBELOWDOMAIN = 0x0002;  // For starting the Find Computer
            private uint BIF_STATUSTEXT = 0x0004;  // Top of the dialog has 2 lines of text for BROWSEINFO.lpszTitle and one line if
                                                   // this flag is set.  Passing the message BFFM_SETSTATUSTEXTA to the hwnd can set the
                                                   // rest of the text.  This is not used with BIF_USENEWUI and BROWSEINFO.lpszTitle gets
                                                   // all three lines of text.
            private uint BIF_RETURNFSANCESTORS = 0x0008;
            private uint BIF_EDITBOX = 0x0010;   // Add an editbox to the dialog
            private uint BIF_VALIDATE = 0x0020;   // insist on valid result (or CANCEL)

            private uint BIF_NEWDIALOGSTYLE = 0x0040;   // Use the new dialog layout with the ability to resize
                                                        // Caller needs to call OleInitialize() before using this API
            private uint BIF_USENEWUI = 0x0040 + 0x0010; //(BIF_NEWDIALOGSTYLE | BIF_EDITBOX);

            private uint BIF_BROWSEINCLUDEURLS = 0x0080;   // Allow URLs to be displayed or entered. (Requires BIF_USENEWUI)
            private uint BIF_UAHINT = 0x0100;   // Add a UA hint to the dialog, in place of the edit box. May not be combined with BIF_EDITBOX
            private uint BIF_NONEWFOLDERBUTTON = 0x0200;   // Do not add the "New Folder" button to the dialog.  Only applicable with BIF_NEWDIALOGSTYLE.
            private uint BIF_NOTRANSLATETARGETS = 0x0400;  // don't traverse target as shortcut

            private uint BIF_BROWSEFORCOMPUTER = 0x1000;  // Browsing for Computers.
            private uint BIF_BROWSEFORPRINTER = 0x2000;// Browsing for Printers
            private uint BIF_BROWSEINCLUDEFILES = 0x4000; // Browsing for Everything
            private uint BIF_SHAREABLE = 0x8000;  // sharable resources displayed (remote shares, requires BIF_USENEWUI)

            [DllImport("shell32.dll")]
            static extern IntPtr SHBrowseForFolder(ref BROWSEINFO lpbi);

            // Note that the BROWSEINFO object's pszDisplayName only gives you the name of the folder.
            // To get the actual path, you need to parse the returned PIDL
            [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
            // static extern uint SHGetPathFromIDList(IntPtr pidl, [MarshalAs(UnmanagedType.LPWStr)] 
            //StringBuilder pszPath);
            static extern bool SHGetPathFromIDList(IntPtr pidl, IntPtr pszPath);

            [DllImport("user32.dll", PreserveSig = true)]
            public static extern IntPtr SendMessage(HandleRef hWnd, uint Msg, int wParam, IntPtr lParam);

            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            public static extern IntPtr SendMessage(HandleRef hWnd, int msg, int wParam, string lParam);

            private string _initialPath;

            public delegate int BrowseCallBackProc(IntPtr hwnd, int msg, IntPtr lp, IntPtr wp);

            /// <summary>
            /// 
            /// </summary>
            struct BROWSEINFO
            {
                public IntPtr hwndOwner;
                public IntPtr pidlRoot;
                public string pszDisplayName;
                public string lpszTitle;
                public uint ulFlags;
                public BrowseCallBackProc lpfn;
                public IntPtr lParam;
                public int iImage;
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="hWnd"></param>
            /// <param name="msg"></param>
            /// <param name="lp"></param>
            /// <param name="lpData"></param>
            /// <returns></returns>
            public int OnBrowseEvent(IntPtr hWnd, int msg, IntPtr lp, IntPtr lpData)
            {
                switch (msg)
                {
                    case BFFM_INITIALIZED: // Required to set initialPath
                        {
                            //Win32.SendMessage(new HandleRef(null, hWnd), BFFM_SETSELECTIONA, 1, lpData);
                            // Use BFFM_SETSELECTIONW if passing a Unicode string, i.e. native CLR Strings.
                            SendMessage(new HandleRef(null, hWnd), BFFM_SETSELECTIONW, 1, _initialPath);
                            break;
                        }
                    case BFFM_SELCHANGED:
                        {
                            IntPtr pathPtr = Marshal.AllocHGlobal((int)(260 * Marshal.SystemDefaultCharSize));
                            if (SHGetPathFromIDList(lp, pathPtr))
                                SendMessage(new HandleRef(null, hWnd), BFFM_SETSTATUSTEXTW, 0, pathPtr);
                            Marshal.FreeHGlobal(pathPtr);
                            break;
                        }
                }

                return 0;
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="caption"></param>
            /// <param name="initialPath"></param>
            /// <param name="parentHandle"></param>
            /// <returns></returns>
            public string SelectFolder(string caption, string initialPath, IntPtr parentHandle)
            {
                _initialPath = initialPath;
                StringBuilder sb = new StringBuilder(256);
                IntPtr bufferAddress = Marshal.AllocHGlobal(256); ;
                IntPtr pidl = IntPtr.Zero;
                BROWSEINFO bi = new BROWSEINFO();
                bi.hwndOwner = parentHandle;
                bi.pidlRoot = IntPtr.Zero;
                bi.lpszTitle = caption;
                bi.ulFlags = BIF_NEWDIALOGSTYLE | BIF_SHAREABLE;
                bi.lpfn = new BrowseCallBackProc(OnBrowseEvent);
                bi.lParam = IntPtr.Zero;
                bi.iImage = 0;

                try
                {
                    pidl = SHBrowseForFolder(ref bi);
                    if (true != SHGetPathFromIDList(pidl, bufferAddress))
                    {
                        return null;
                    }
                    sb.Append(Marshal.PtrToStringAuto(bufferAddress));
                }
                finally
                {
                    // Caller is responsible for freeing this memory.
                    Marshal.FreeCoTaskMem(pidl);
                }

                return sb.ToString();
            }
        }
    }
}

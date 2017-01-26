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

// Disable warning CS0108: 'x' hides inherited member 'y'. Use the new keyword if hiding was intended.
#pragma warning disable 0108

namespace Nequeo.Interop
{
    /// <summary>
    /// Open and Save COM dialog interface.
    /// </summary>
    static class FileDialogNative
    {
        [ComImport]
        [ClassInterface(ClassInterfaceType.None)]
        [TypeLibType(TypeLibTypeFlags.FCanCreate)]
        [Guid(CLSIDGuid.FileOpenDialog)]
        internal class FileOpenDialogRCW
        { }

        [ComImport]
        [ClassInterface(ClassInterfaceType.None)]
        [TypeLibType(TypeLibTypeFlags.FCanCreate)]
        [Guid(CLSIDGuid.FileSaveDialog)]
        internal class FileSaveDialogRCW
        { }

        [ComImport]
        [Guid(IIDGuid.IFileOpenDialog)]
        [CoClass(typeof(FileOpenDialogRCW))]
        internal interface NativeFileOpenDialog : IFileOpenDialog
        { }

        [ComImport]
        [Guid(IIDGuid.IFileSaveDialog)]
        [CoClass(typeof(FileSaveDialogRCW))]
        internal interface NativeFileSaveDialog : IFileSaveDialog
        { }

        /// <summary>
        /// 
        /// </summary>
        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        internal struct PROPERTYKEY
        {
            internal Guid fmtid;
            internal uint pid;
        }

        /// <summary>
        /// 
        /// </summary>
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto, Pack = 4)]
        internal struct COMDLG_FILTERSPEC
        {
            [MarshalAs(UnmanagedType.LPWStr)]
            internal string pszName;

            [MarshalAs(UnmanagedType.LPWStr)]
            internal string pszSpec;
        }

        /// <summary>
        /// 
        /// </summary>
        internal class IIDGuid
        {
            private IIDGuid() { } // Avoid FxCop violation AvoidUninstantiatedInternalClasses
            internal const string IModalWindow = "b4db1657-70d7-485e-8e3e-6fcb5a5c1802";
            internal const string IFileDialog = "42f85136-db7e-439c-85f1-e4075d135fc8";
            internal const string IFileOpenDialog = "d57c7288-d4ad-4768-be02-9d969532d960";
            internal const string IFileSaveDialog = "84bccd23-5fde-4cdb-aea4-af64b83d78ab";
            internal const string IFileDialogEvents = "973510DB-7D7F-452B-8975-74A85828D354";
            internal const string IShellItem = "43826D1E-E718-42EE-BC55-A1E261C37BFE";
            internal const string IShellItemArray = "B63EA76D-1F85-456F-A19C-48159EFA858B";
            internal const string IShellItemFilter = "2659B475-EEB8-48b7-8F07-B378810F48CF";

        }

        /// <summary>
        /// 
        /// </summary>
        internal class CLSIDGuid
        {
            private CLSIDGuid() { } // Avoid FxCop violation AvoidUninstantiatedInternalClasses
            internal const string FileOpenDialog = "DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7";
            internal const string FileSaveDialog = "C0B4E2F3-BA21-4773-8DBA-335EC946EB8B";
        }

        /// <summary>
        /// 
        /// </summary>
        [Flags]
        internal enum FOS : uint
        {
            FOS_OVERWRITEPROMPT = 0x00000002,
            FOS_STRICTFILETYPES = 0x00000004,
            FOS_NOCHANGEDIR = 0x00000008,
            FOS_PICKFOLDERS = 0x00000020,
            FOS_FORCEFILESYSTEM = 0x00000040,
            FOS_ALLNONSTORAGEITEMS = 0x00000080,
            FOS_NOVALIDATE = 0x00000100,
            FOS_ALLOWMULTISELECT = 0x00000200,
            FOS_PATHMUSTEXIST = 0x00000800,
            FOS_FILEMUSTEXIST = 0x00001000,
            FOS_CREATEPROMPT = 0x00002000,
            FOS_SHAREAWARE = 0x00004000,
            FOS_NOREADONLYRETURN = 0x00008000,
            FOS_NOTESTFILECREATE = 0x00010000,
            FOS_HIDEMRUPLACES = 0x00020000,
            FOS_HIDEPINNEDPLACES = 0x00040000,
            FOS_NODEREFERENCELINKS = 0x00100000,
            FOS_DONTADDTORECENT = 0x02000000,
            FOS_FORCESHOWHIDDEN = 0x10000000,
            FOS_DEFAULTNOMINIMODE = 0x20000000,
            FOS_FORCEPREVIEWPANEON = 0x40000000,
            FOS_SUPPORTSTREAMABLEITEMS = 0x80000000
        }

        /// <summary>
        /// 
        /// </summary>
        internal enum FDAP : uint
        {
            FDAP_BOTTOM = 0,
            FDAP_TOP = 1
        }

        /// <summary>
        /// 
        /// </summary>
        [Flags]
        internal enum SIATTRIBFLAGS
        {
            SIATTRIBFLAGS_AND = 0x00000001,
            SIATTRIBFLAGS_OR = 0x00000002,
            SIATTRIBFLAGS_APPCOMPAT = 0x00000003,
            SIATTRIBFLAGS_MASK = 0x00000003,
            SIATTRIBFLAGS_ALLITEMS = 0x00004000,
        }

        /// <summary>
        /// 
        /// </summary>
        internal enum FDE_OVERWRITE_RESPONSE
        {
            FDEOR_DEFAULT = 0,
            FDEOR_ACCEPT = 1,
            FDEOR_REFUSE = 2
        }

        /// <summary>
        /// 
        /// </summary>
        internal enum FDE_SHAREVIOLATION_RESPONSE
        {
            FDESVR_DEFAULT = 0,
            FDESVR_ACCEPT = 1,
            FDESVR_REFUSE = 2
        }

        /// <summary>
        /// 
        /// </summary>
        internal enum SIGDN : uint
        {
            SIGDN_NORMALDISPLAY = 0x00000000,               // SHGDN_NORMAL
            SIGDN_PARENTRELATIVEPARSING = 0x80018001,       // SHGDN_INFOLDER | SHGDN_FORPARSING
            SIGDN_DESKTOPABSOLUTEPARSING = 0x80028000,      // SHGDN_FORPARSING
            SIGDN_PARENTRELATIVEEDITING = 0x80031001,       // SHGDN_INFOLDER | SHGDN_FOREDITING
            SIGDN_DESKTOPABSOLUTEEDITING = 0x8004c000,      // SHGDN_FORPARSING | SHGDN_FORADDRESSBAR
            SIGDN_FILESYSPATH = 0x80058000,                 // SHGDN_FORPARSING
            SIGDN_URL = 0x80068000,                         // SHGDN_FORPARSING
            SIGDN_PARENTRELATIVEFORADDRESSBAR = 0x8007c001, // SHGDN_INFOLDER | SHGDN_FORPARSING | SHGDN_FORADDRESSBAR
            SIGDN_PARENTRELATIVE = 0x80080001,              // SHGDN_INFOLDER
            SIGDN_PARENTRELATIVEFORUI = 0x80094001
        }

        /// <summary>
        /// 
        /// </summary>
        internal enum HRESULT : long
        {
            S_FALSE = 0x0001,
            S_OK = 0x0000,
            E_INVALIDARG = 0x80070057,
            E_OUTOFMEMORY = 0x8007000E
        }

        /// <summary>
        /// 
        /// </summary>
        [Flags]
        internal enum SHCONTF : long
        {
            SHCONTF_CHECKING_FOR_CHILDREN = 0x00010,
            SHCONTF_FOLDERS = 0x00020,
            SHCONTF_NONFOLDERS = 0x00040,
            SHCONTF_INCLUDEHIDDEN = 0x00080,
            SHCONTF_INIT_ON_FIRST_NEXT = 0x00100,
            SHCONTF_NETPRINTERSRCH = 0x00200,
            SHCONTF_SHAREABLE = 0x00400,
            SHCONTF_STORAGE = 0x00800,
            SHCONTF_NAVIGATION_ENUM = 0x01000,
            SHCONTF_FASTITEMS = 0x02000,
            SHCONTF_FLATLIST = 0x04000,
            SHCONTF_ENABLE_ASYNC = 0x08000,
            SHCONTF_INCLUDESUPERHIDDEN = 0x10000
        }

        /// <summary>
        /// 
        /// </summary>
        [ComImport()]
        [Guid(IIDGuid.IModalWindow)]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IModalWindow
        {
            [PreserveSig]
            int Show([In] IntPtr parent);
        }

        /// <summary>
        /// 
        /// </summary>
        [ComImport(),
        Guid(IIDGuid.IFileDialog),
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IFileDialog
        {
            [PreserveSig]
            int Show([In] IntPtr parent);

            void SetFileTypes([In] uint cFileTypes, [In] [MarshalAs(UnmanagedType.LPArray)]COMDLG_FILTERSPEC[] rgFilterSpec);

            void SetFileTypeIndex([In] uint iFileType);

            void GetFileTypeIndex(out uint piFileType);

            void Advise([In, MarshalAs(UnmanagedType.Interface)] IFileDialogEvents pfde, out uint pdwCookie);

            void Unadvise([In] uint dwCookie);

            void SetOptions([In] FOS fos);

            void GetOptions(out FOS pfos);

            void SetDefaultFolder([In, MarshalAs(UnmanagedType.Interface)] IShellItem psi);

            void SetFolder([In, MarshalAs(UnmanagedType.Interface)] IShellItem psi);

            void GetFolder([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);

            void GetCurrentSelection([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);

            void SetFileName([In, MarshalAs(UnmanagedType.LPWStr)] string pszName);

            void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);

            void SetTitle([In, MarshalAs(UnmanagedType.LPWStr)] string pszTitle);

            void SetOkButtonLabel([In, MarshalAs(UnmanagedType.LPWStr)] string pszText);

            void SetFileNameLabel([In, MarshalAs(UnmanagedType.LPWStr)] string pszLabel);

            void GetResult([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);

            void AddPlace([In, MarshalAs(UnmanagedType.Interface)] IShellItem psi, int alignment);

            void SetDefaultExtension([In, MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);

            void Close([MarshalAs(UnmanagedType.Error)] int hr);

            void SetClientGuid([In] ref Guid guid);

            void ClearClientData();

            // Not supported:  IShellItemFilter is not defined, converting to IntPtr
            void SetFilter([MarshalAs(UnmanagedType.Interface)] IShellItemFilter pFilter);
        }

        /// <summary>
        /// 
        /// </summary>
        [ComImport(),
        Guid(IIDGuid.IFileOpenDialog),
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IFileOpenDialog : IFileDialog
        {
            [PreserveSig]
            int Show([In] IntPtr parent);

            void SetFileTypes([In] uint cFileTypes, [In] ref COMDLG_FILTERSPEC rgFilterSpec);

            void SetFileTypeIndex([In] uint iFileType);

            void GetFileTypeIndex(out uint piFileType);

            void Advise([In, MarshalAs(UnmanagedType.Interface)] IFileDialogEvents pfde, out uint pdwCookie);

            void Unadvise([In] uint dwCookie);

            void SetOptions([In] FOS fos);

            void GetOptions(out FOS pfos);

            void SetDefaultFolder([In, MarshalAs(UnmanagedType.Interface)] IShellItem psi);

            void SetFolder([In, MarshalAs(UnmanagedType.Interface)] IShellItem psi);

            void GetFolder([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);

            void GetCurrentSelection([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);

            void SetFileName([In, MarshalAs(UnmanagedType.LPWStr)] string pszName);

            void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);

            void SetTitle([In, MarshalAs(UnmanagedType.LPWStr)] string pszTitle);

            void SetOkButtonLabel([In, MarshalAs(UnmanagedType.LPWStr)] string pszText);

            void SetFileNameLabel([In, MarshalAs(UnmanagedType.LPWStr)] string pszLabel);

            void GetResult([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);

            void SetDefaultExtension([In, MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);

            void Close([MarshalAs(UnmanagedType.Error)] int hr);

            void SetClientGuid([In] ref Guid guid);

            void ClearClientData();

            void SetFilter([MarshalAs(UnmanagedType.Interface)] IntPtr pFilter);

            void GetResults([MarshalAs(UnmanagedType.Interface)] out IShellItemArray ppenum);

            void GetSelectedItems([MarshalAs(UnmanagedType.Interface)] out IShellItemArray ppsai);
        }

        /// <summary>
        /// 
        /// </summary>
        [ComImport(),
        Guid(IIDGuid.IFileSaveDialog),
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IFileSaveDialog : IFileDialog
        {
            [PreserveSig]
            int Show([In] IntPtr parent);

            void SetFileTypes([In] uint cFileTypes, [In] ref COMDLG_FILTERSPEC rgFilterSpec);

            void SetFileTypeIndex([In] uint iFileType);

            void GetFileTypeIndex(out uint piFileType);

            void Advise([In, MarshalAs(UnmanagedType.Interface)] IFileDialogEvents pfde, out uint pdwCookie);

            void Unadvise([In] uint dwCookie);

            void SetOptions([In] FOS fos);

            void GetOptions(out FOS pfos);

            void SetDefaultFolder([In, MarshalAs(UnmanagedType.Interface)] IShellItem psi);

            void SetFolder([In, MarshalAs(UnmanagedType.Interface)] IShellItem psi);

            void GetFolder([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);

            void GetCurrentSelection([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);

            void SetFileName([In, MarshalAs(UnmanagedType.LPWStr)] string pszName);

            void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);

            void SetTitle([In, MarshalAs(UnmanagedType.LPWStr)] string pszTitle);

            void SetOkButtonLabel([In, MarshalAs(UnmanagedType.LPWStr)] string pszText);

            void SetFileNameLabel([In, MarshalAs(UnmanagedType.LPWStr)] string pszLabel);

            void GetResult([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);

            void SetDefaultExtension([In, MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);

            void Close([MarshalAs(UnmanagedType.Error)] int hr);

            void SetClientGuid([In] ref Guid guid);

            void ClearClientData();

            void SetFilter([MarshalAs(UnmanagedType.Interface)] IntPtr pFilter);

            void SetSaveAsItem([In, MarshalAs(UnmanagedType.Interface)] IShellItem psi);

            void SetProperties([In, MarshalAs(UnmanagedType.Interface)] IntPtr pStore);

            void SetCollectedProperties([In, MarshalAs(UnmanagedType.Interface)] IntPtr pList, [In] int fAppendDefault);

            void GetProperties([MarshalAs(UnmanagedType.Interface)] out IntPtr ppStore);

            void ApplyProperties(
                [In, MarshalAs(UnmanagedType.Interface)] IShellItem psi,
                [In, MarshalAs(UnmanagedType.Interface)] IntPtr pStore,
                [In, ComAliasName("ShellObjects.wireHWND")] ref IntPtr hwnd,
                [In, MarshalAs(UnmanagedType.Interface)] IntPtr pSink);
        }

        /// <summary>
        /// 
        /// </summary>
        [ComImport]
        [Guid(IIDGuid.IShellItemArray)]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IShellItemArray
        {
            // Not supported: IBindCtx
            void BindToHandler([In, MarshalAs(UnmanagedType.Interface)] IntPtr pbc, [In] ref Guid rbhid, [In] ref Guid riid, out IntPtr ppvOut);

            void GetPropertyStore([In] int Flags, [In] ref Guid riid, out IntPtr ppv);

            void GetPropertyDescriptionList([In] ref PROPERTYKEY keyType, [In] ref Guid riid, out IntPtr ppv);

            void GetAttributes([In] SIATTRIBFLAGS dwAttribFlags, [In] uint sfgaoMask, out uint psfgaoAttribs);

            void GetCount(out uint pdwNumItems);

            void GetItemAt([In] uint dwIndex, [MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);

            // Not supported: IEnumShellItems (will use GetCount and GetItemAt instead)
            void EnumItems([MarshalAs(UnmanagedType.Interface)] out IntPtr ppenumShellItems);
        }

        /// <summary>
        /// 
        /// </summary>
        [ComImport,
        Guid(IIDGuid.IShellItem),
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IShellItem
        {
            void BindToHandler(IntPtr pbc,
                [MarshalAs(UnmanagedType.LPStruct)]Guid bhid,
                [MarshalAs(UnmanagedType.LPStruct)]Guid riid,
                out IntPtr ppv);

            void GetParent(out IShellItem ppsi);

            void GetDisplayName(SIGDN sigdnName, out IntPtr ppszName);

            void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);

            void Compare(IShellItem psi, uint hint, out int piOrder);
        }

        /// <summary>
        /// 
        /// </summary>
        [ComImport,
        Guid(IIDGuid.IShellItemFilter),
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IShellItemFilter
        {
            int IncludeItem([In, MarshalAs(UnmanagedType.Interface)] IShellItem psi);

            [PreserveSig]
            int GetEnumFlagsForItem([In, MarshalAs(UnmanagedType.Interface)] IShellItem psi, out SHCONTF pgrfFlags);
        }

        /// <summary>
        /// 
        /// </summary>
        [ComImport,
        Guid(IIDGuid.IFileDialogEvents),
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IFileDialogEvents
        {
            [PreserveSig]
            HRESULT OnFileOk([In, MarshalAs(UnmanagedType.Interface)] IFileDialog pfd);

            [PreserveSig]
            HRESULT OnFolderChanging([In, MarshalAs(UnmanagedType.Interface)] IFileDialog pfd,
                           [In, MarshalAs(UnmanagedType.Interface)] IShellItem psiFolder);

            void OnFolderChange([In, MarshalAs(UnmanagedType.Interface)] IFileDialog pfd);

            void OnSelectionChange([In, MarshalAs(UnmanagedType.Interface)] IFileDialog pfd);

            void OnShareViolation([In, MarshalAs(UnmanagedType.Interface)] IFileDialog pfd,
                        [In, MarshalAs(UnmanagedType.Interface)] IShellItem psi,
                        out FDE_SHAREVIOLATION_RESPONSE pResponse);

            void OnTypeChange([In, MarshalAs(UnmanagedType.Interface)] IFileDialog pfd);

            void OnOverwrite([In, MarshalAs(UnmanagedType.Interface)] IFileDialog pfd,
                       [In, MarshalAs(UnmanagedType.Interface)] IShellItem psi,
                       out FDE_OVERWRITE_RESPONSE pResponse);
        }
    }
}

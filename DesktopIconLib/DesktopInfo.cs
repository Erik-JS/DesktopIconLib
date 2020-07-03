using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using static DesktopIconLib.NativeMethods;

namespace DesktopIconLib
{
    public class DesktopInfo
    {

        private static IntPtr handleListView = IntPtr.Zero;

        private static bool Is64bit = Environment.Is64BitOperatingSystem;

        private static void UpdateListViewHandle()
        {
            IntPtr handleProgman, handleShelldll;
            handleProgman = FindWindow("Progman", null);
            handleShelldll = FindWindowEx(handleProgman, IntPtr.Zero, "SHELLDLL_DefView", null);
            if (handleShelldll == IntPtr.Zero)
            {
                EnumWindows(delegate (IntPtr hWnd, IntPtr lParam) {
                    if (GetClassNameString(hWnd) == "WorkerW")
                    {
                        handleShelldll = FindWindowEx(hWnd, IntPtr.Zero, "SHELLDLL_DefView", null);
                        if (handleShelldll != IntPtr.Zero)
                        {
                            return false;
                        }
                    }
                    return true;
                }
                , IntPtr.Zero);
            }
            handleListView = FindWindowEx(handleShelldll, IntPtr.Zero, "SysListView32", null);
        }

        public static int GetIconCount()
        {
            UpdateListViewHandle();
            IntPtr iconccount = SendMessage(handleListView, LVM_GETITEMCOUNT, 0, IntPtr.Zero);
            return iconccount.ToInt32();
        }

        private static string GetClassNameString(IntPtr hWnd)
        {
            StringBuilder sb = new StringBuilder(128);
            int a = GetClassName(hWnd, sb, 128);
            if (a == 0) return "";
            return sb.ToString();
        }

        public static List<string> GetIconLabels()
        {
            List<string> lstItems = new List<string>();
            int MaxChar = 0x100;
            int itemCount = GetIconCount();
            GetWindowThreadProcessId(handleListView, out uint pid);
            IntPtr handleX = OpenProcess(ProcessAccessFlags.All, false, pid);
            if (handleX == IntPtr.Zero)
            {
                System.Diagnostics.Debug.WriteLine("*** OpenProcess failed ***");
                return lstItems;
            }

            IntPtr memLoc = VirtualAllocEx(handleX, IntPtr.Zero, 0x1000, AllocationType.Commit, MemoryProtection.ReadWrite);
            object lvItem;
            if (Is64bit)
                lvItem = new LVITEMA64() { mask = LVIF_TEXT, iSubItem = 0, cchTextMax = MaxChar };
            else
                lvItem = new LVITEMA32() { mask = LVIF_TEXT, iSubItem = 0, cchTextMax = MaxChar };

            int lvItemSize = Marshal.SizeOf(lvItem);
            byte[] itemBuffer = new byte[lvItemSize];
            LVITEMA32 lvItem32;
            LVITEMA64 lvItem64;
            for (int i = 0; i < itemCount; i++)
            {
                if(Is64bit)
                {
                    lvItem64 = (LVITEMA64)lvItem;
                    lvItem64.iItem = i;
                    lvItem64.pszText = (long)(memLoc + 0x300);
                    lvItem = lvItem64;
                }
                else
                {
                    lvItem32 = (LVITEMA32)lvItem;
                    lvItem32.iItem = i;
                    lvItem32.pszText = (int)(memLoc + 0x300);
                    lvItem = lvItem32;
                }
                // alloc mem for unmanaged obj
                var lvItemLocalPtr = Marshal.AllocHGlobal(lvItemSize);
                // copy struct to unmanaged space
                Marshal.StructureToPtr(lvItem, lvItemLocalPtr, false);
                // copy unmanaged struct to the target process
                WriteProcessMemory(handleX, memLoc, lvItemLocalPtr, (uint)lvItemSize, IntPtr.Zero);

                SendMessage(handleListView, LVM_GETITEMW, i, memLoc);

                // read updated item from target processs
                ReadProcessMemory(handleX, memLoc, Marshal.UnsafeAddrOfPinnedArrayElement(itemBuffer, 0), (uint)lvItemSize, IntPtr.Zero);
                IntPtr strmemloc;
                if(Is64bit)
                {
                    lvItem64 = (LVITEMA64)Marshal.PtrToStructure(Marshal.UnsafeAddrOfPinnedArrayElement(itemBuffer, 0), typeof(LVITEMA64));
                    strmemloc = (IntPtr)lvItem64.pszText;
                }
                else
                {
                    lvItem32 = (LVITEMA32)Marshal.PtrToStructure(Marshal.UnsafeAddrOfPinnedArrayElement(itemBuffer, 0), typeof(LVITEMA32));
                    strmemloc = (IntPtr)lvItem32.pszText;
                }
                string str = ReadString(handleX, strmemloc, MaxChar);
                lstItems.Add(str);

                Marshal.FreeHGlobal(lvItemLocalPtr);
            }

            VirtualFreeEx(handleX, memLoc, 0, AllocationType.Release);
            CloseHandle(handleX);
            return lstItems;
        }

        private static string ReadString(IntPtr handle, IntPtr loc, int maxcharcount)
        {
            byte[] bytebuffer = new byte[maxcharcount * 2];
            ReadProcessMemory(handle, loc, bytebuffer, (uint)bytebuffer.Length, IntPtr.Zero);
            string str = System.Text.Encoding.Unicode.GetString(bytebuffer);
            str = str.Substring(0, str.IndexOf('\0'));
            return str;
        }

        public static List<Tuple<int, int>> GetIconCoordinates()
        {
            var lstPositions = new List<Tuple<int, int>>();
            int itemCount = GetIconCount();
            GetWindowThreadProcessId(handleListView, out uint pid);
            IntPtr handleX = OpenProcess(ProcessAccessFlags.All, false, pid);
            if (handleX == IntPtr.Zero)
            {
                System.Diagnostics.Debug.WriteLine("*** OpenProcess failed ***");
                return lstPositions;
            }

            IntPtr memLoc = VirtualAllocEx(handleX, IntPtr.Zero, 0x1000, AllocationType.Commit, MemoryProtection.ReadWrite);
            POINT pp = new POINT();
            int pointVarSize = Marshal.SizeOf(pp);
            byte[] pBuffer = new byte[pointVarSize];

            for (int i = 0; i < itemCount; i++)
            {
                // alloc mem for unmanaged obj
                var pointUnmanagedPtr = Marshal.AllocHGlobal(pointVarSize);
                // copy struct to unmanaged space
                Marshal.StructureToPtr(pp, pointUnmanagedPtr, false);
                // copy unmanaged struct to the target process
                WriteProcessMemory(handleX, memLoc, pointUnmanagedPtr, (uint)pointVarSize, IntPtr.Zero);

                SendMessage(handleListView, LVM_GETITEMPOSITION, i, memLoc);

                ReadProcessMemory(handleX, memLoc, Marshal.UnsafeAddrOfPinnedArrayElement(pBuffer, 0), (uint)pointVarSize, IntPtr.Zero);
                pp = (POINT)Marshal.PtrToStructure(Marshal.UnsafeAddrOfPinnedArrayElement(pBuffer, 0), typeof(POINT));
                lstPositions.Add(new Tuple<int, int>(pp.X, pp.Y));

                Marshal.FreeHGlobal(pointUnmanagedPtr);
            }

            VirtualFreeEx(handleX, memLoc, 0, AllocationType.Release);
            CloseHandle(handleX);
            return lstPositions;
        }

        public static void SetIconCoordinates(List<string> labels, List<Tuple<int,int>> coords)
        {
            UpdateListViewHandle();
            GetWindowThreadProcessId(handleListView, out uint pid);
            IntPtr handleX = OpenProcess(ProcessAccessFlags.All, false, pid);
            IntPtr memLoc = VirtualAllocEx(handleX, IntPtr.Zero, 0x1000, AllocationType.Commit, MemoryProtection.ReadWrite);
            object lvFindInfo;
            if (Is64bit)
                lvFindInfo = new LVFINDINFOW64();
            else
                lvFindInfo = new LVFINDINFOW32();
            
            int lvfindSize = Marshal.SizeOf(lvFindInfo);

            for (int i = 0; i < labels.Count; i++)
            {
                if(Is64bit)
                {
                    LVFINDINFOW64 lvFindInfo64 = (LVFINDINFOW64)lvFindInfo;
                    lvFindInfo64.flags = LVFI_STRING;
                    lvFindInfo64.psz = (long)(memLoc + 0x300);
                    lvFindInfo = lvFindInfo64;
                }
                else
                {
                    LVFINDINFOW32 lvFindInfo32 = (LVFINDINFOW32)lvFindInfo;
                    lvFindInfo32.flags = LVFI_STRING;
                    lvFindInfo32.psz = (int)(memLoc + 0x300);
                    lvFindInfo = lvFindInfo32;
                }
                // alloc mem for unmanaged obj
                var lvfiUnmanagedPtr = Marshal.AllocHGlobal(lvfindSize);
                // copy struct to unmanaged space
                Marshal.StructureToPtr(lvFindInfo, lvfiUnmanagedPtr, false);
                // copy unmanaged struct to the target process
                WriteProcessMemory(handleX, memLoc, lvfiUnmanagedPtr, (uint)lvfindSize, IntPtr.Zero);
                WriteString(handleX, memLoc + 0x300, labels[i]);

                int idx = (int)SendMessage(handleListView, LVM_FINDITEMW, -1, memLoc);
                if(idx != -1)
                {
                    POINT pp = new POINT() { X = coords[i].Item1, Y = coords[i].Item2 };
                    int pointSize = Marshal.SizeOf(pp);
                    var pointUnmanagedPtr = Marshal.AllocHGlobal(pointSize);
                    Marshal.StructureToPtr(pp, pointUnmanagedPtr, false);
                    WriteProcessMemory(handleX, memLoc + 0x800, pointUnmanagedPtr, (uint)pointSize, IntPtr.Zero);
                    SendMessage(handleListView, LVM_SETITEMPOSITION32, idx, memLoc + 0x800);
                    Marshal.FreeHGlobal(pointUnmanagedPtr);
                }
                Marshal.FreeHGlobal(lvfiUnmanagedPtr);
            }
            VirtualFreeEx(handleX, memLoc, 0, AllocationType.Release);
            CloseHandle(handleX);
        }

        public static void SetIconCoordinates(string label, int X, int Y)
        {
            SetIconCoordinates(new List<string>() { label }, new List<Tuple<int, int>>() { new Tuple<int, int>(X, Y) });
        }

        public static void SetIconCoordinates(string[] labelArray, int[] xArray, int[] yArray)
        {
            var lstCoords = new List<Tuple<int, int>>();
            for (int i = 0; i < xArray.Length; i++)
            {
                lstCoords.Add(new Tuple<int, int>(xArray[i], yArray[i]));
            }
            SetIconCoordinates(new List<string>(labelArray), lstCoords);
        }

        private static void WriteString(IntPtr handle, IntPtr loc, string text)
        {
            var lstBytes = System.Text.Encoding.Unicode.GetBytes(text).ToList();
            lstBytes.AddRange(new byte[] { 0, 0 });
            var arrayBytes = lstBytes.ToArray();
            WriteProcessMemory(handle, loc, arrayBytes, (uint)arrayBytes.Length, IntPtr.Zero);
        }

        public static List<DesktopIcon> GetListOfIcons()
        {
            var res = new List<DesktopIcon>();
            var listNames = GetIconLabels();
            var listCoords = GetIconCoordinates();
            for (int i = 0; i < listNames.Count; i++)
            {
                res.Add(new DesktopIcon(listNames[i], listCoords[i].Item1, listCoords[i].Item2));
            }
            return res;
        }

    }

}
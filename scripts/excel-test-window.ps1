Add-Type @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public class ExcelTestWindow {
 public delegate bool Callback(IntPtr h, IntPtr l);
 [DllImport("user32.dll")] public static extern bool EnumWindows(Callback cb, IntPtr l);
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
 [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder text, int length);
 [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
 public static void DismissReminder(uint process) {
   EnumWindows((h,l) => { uint pid; GetWindowThreadProcessId(h,out pid); var s=new StringBuilder(200); GetWindowText(h,s,200);
     if(pid==process && s.ToString()=="Microsoft Office-Aktivierungs-Assistent") PostMessage(h,16,IntPtr.Zero,IntPtr.Zero);
     return true; },IntPtr.Zero);
 }
}
'@

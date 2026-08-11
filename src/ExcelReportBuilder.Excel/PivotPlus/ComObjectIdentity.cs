using System;
using System.Runtime.InteropServices;

namespace ExcelReportBuilder.Excel.PivotPlus
{
    internal static class ComObjectIdentity
    {
        public static bool AreSame(object? left, object? right)
        {
            if (left == null || right == null) return false;
            if (ReferenceEquals(left, right)) return true;
            if (!Marshal.IsComObject(left) || !Marshal.IsComObject(right)) return false;

            IntPtr leftIdentity = IntPtr.Zero;
            IntPtr rightIdentity = IntPtr.Zero;
            try
            {
                leftIdentity = Marshal.GetIUnknownForObject(left);
                rightIdentity = Marshal.GetIUnknownForObject(right);
                return leftIdentity == rightIdentity;
            }
            finally
            {
                if (leftIdentity != IntPtr.Zero) Marshal.Release(leftIdentity);
                if (rightIdentity != IntPtr.Zero) Marshal.Release(rightIdentity);
            }
        }
    }
}

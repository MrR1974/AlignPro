using System.Runtime.InteropServices;

namespace AlignPro.AddIn
{
    /// <summary>
    /// COM housekeeping. Every interop object we take a reference to has to be released, or a long
    /// editing session slowly destabilises PowerPoint - which is also why nothing in this add-in uses
    /// chained COM expressions: they leave intermediate references we never get the chance to free.
    /// </summary>
    internal static class Com
    {
        /// <summary>
        /// Releases an interop object if it is one. Safe to call with null, and safe to call on a
        /// managed object, so it can be used unconditionally in a finally block.
        /// </summary>
        public static void Release(object? comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }
    }
}

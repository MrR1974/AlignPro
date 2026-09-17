using System;
using System.Globalization;
using System.IO;

namespace AlignPro.AddIn
{
    /// <summary>
    /// Append-only trace log, written to <c>%LOCALAPPDATA%\AlignPro\alignpro.log</c>.
    /// </summary>
    /// <remarks>
    /// An add-in is awkward to debug: it lives inside PowerPoint, its ribbon callbacks are invoked by
    /// the host through IDispatch, and a callback that is never resolved fails silently and looks
    /// exactly like a callback that ran and decided to do nothing. A log distinguishes the two, which
    /// is the whole reason this exists.
    ///
    /// Every write is guarded. Logging must never be the thing that breaks the add-in.
    /// </remarks>
    internal static class Diagnostics
    {
        private static readonly object Gate = new object();
        private static string? _path;

        /// <summary>Turn tracing off to stop all file access.</summary>
        public static bool Enabled { get; set; } = true;

        public static string LogPath => _path ??= BuildPath();

        private static string BuildPath()
        {
            var directory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AlignPro");
            return Path.Combine(directory, "alignpro.log");
        }

        public static void Log(string message)
        {
            if (!Enabled) return;

            try
            {
                lock (Gate)
                {
                    var directory = Path.GetDirectoryName(LogPath);
                    if (directory != null && !Directory.Exists(directory)) Directory.CreateDirectory(directory);

                    var line = string.Format(
                        CultureInfo.InvariantCulture,
                        "{0:yyyy-MM-dd HH:mm:ss.fff}  {1}{2}",
                        DateTime.Now, message, Environment.NewLine);

                    File.AppendAllText(LogPath, line);
                }
            }
            catch
            {
                // A failed log write is never worth surfacing to the user.
            }
        }
    }
}

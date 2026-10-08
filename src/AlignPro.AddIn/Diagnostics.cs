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

        /// <summary>
        /// Per-shape and per-change detail. Off by default because it is noisy, but it is what
        /// identified a grid bug by showing exactly what the reader read and the solver produced.
        /// </summary>
        public static bool Verbose { get; set; }

        public static string LogPath => _path ??= BuildPath();

        private static string BuildPath()
        {
            var directory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AlignPro");
            return Path.Combine(directory, "alignpro.log");
        }

        /// <summary>Starts timing the stages of one command, for the line it logs when done.</summary>
        public static StageTimer StartStages() => new StageTimer();

        /// <summary>Detail that is only written when <see cref="Verbose"/> is on.</summary>
        public static void LogVerbose(string message)
        {
            if (Verbose) Log(message);
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

    /// <summary>
    /// How long each stage of a command took - reading the selection, solving, writing back - so
    /// the log shows where the time went, not just that a command was slow.
    /// </summary>
    internal sealed class StageTimer
    {
        private readonly System.Diagnostics.Stopwatch _watch = System.Diagnostics.Stopwatch.StartNew();
        private readonly System.Text.StringBuilder _stages = new System.Text.StringBuilder();
        private long _last;

        /// <summary>Records the time since the previous mark under <paramref name="stage"/>.</summary>
        public void Mark(string stage)
        {
            var now = _watch.ElapsedMilliseconds;
            if (_stages.Length > 0) _stages.Append(' ');
            _stages.Append(stage).Append('=').Append((now - _last).ToString(CultureInfo.InvariantCulture)).Append("ms");
            _last = now;
        }

        public override string ToString() => _stages.ToString();
    }
}

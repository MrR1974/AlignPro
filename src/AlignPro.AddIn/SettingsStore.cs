using System;
using System.Collections.Generic;
using Microsoft.Win32;

namespace AlignPro.AddIn
{
    /// <summary>
    /// Remembers the ribbon's settings between PowerPoint sessions, one string value per setting under
    /// <c>HKEY_CURRENT_USER\Software\AlignPro\Settings</c>.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The registry rather than a file because each value is written on its own: a crash mid-write
    /// cannot leave a half-written file that loses every setting at once. It is also per-user, like the
    /// install, and outlives upgrades.
    /// </para>
    /// <para>
    /// Not <c>Properties.Settings</c>: in a VSTO add-in that stores <c>user.config</c> under a folder
    /// named for the install location and version, so every upgrade would start from scratch and no
    /// uninstaller could find it.
    /// </para>
    /// <para>
    /// Every value is written as it changes, because PowerPoint does not reliably raise Shutdown for
    /// COM add-ins - there is no dependable moment to save on the way out. Every access is guarded:
    /// a registry that cannot be read means the defaults, never a broken ribbon.
    /// </para>
    /// </remarks>
    internal sealed class SettingsStore
    {
        /// <summary>
        /// Under HKEY_CURRENT_USER. The uninstallers delete <c>Software\AlignPro</c> as a whole, so keep
        /// anything else AlignPro ever stores beneath it too.
        /// </summary>
        internal const string KeyPath = @"Software\AlignPro\Settings";

        /// <summary>Every stored value by name. Empty when nothing is stored or the key cannot be read.</summary>
        public IReadOnlyDictionary<string, string> ReadAll()
        {
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(KeyPath))
                {
                    if (key == null) return values;

                    foreach (var name in key.GetValueNames())
                    {
                        if (key.GetValue(name) is string value) values[name] = value;
                    }
                }
            }
            catch (Exception ex)
            {
                Diagnostics.Log("Could not read saved settings, so using the defaults: " + ex.Message);
            }

            return values;
        }

        public void Write(string name, string value)
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(KeyPath))
                {
                    key?.SetValue(name, value, RegistryValueKind.String);
                }
            }
            catch (Exception ex)
            {
                Diagnostics.Log("Could not save setting " + name + ": " + ex.Message);
            }
        }

        /// <summary>Forgets every saved setting, so the next session starts from the defaults.</summary>
        public void Clear()
        {
            try
            {
                Registry.CurrentUser.DeleteSubKeyTree(KeyPath, throwOnMissingSubKey: false);
            }
            catch (Exception ex)
            {
                Diagnostics.Log("Could not clear saved settings: " + ex.Message);
            }
        }
    }
}

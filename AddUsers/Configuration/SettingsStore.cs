using System;
using System.IO;
using Newtonsoft.Json;

namespace AddUsers.Configuration
{
    /// <summary>
    /// Loads and saves <see cref="AddInSettings"/> as JSON under %APPDATA%\AddUsers.
    /// </summary>
    public static class SettingsStore
    {
        private static readonly object SyncRoot = new object();

        public static string SettingsDirectory =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AddUsers");

        public static string SettingsPath => Path.Combine(SettingsDirectory, "settings.json");

        /// <summary>
        /// Loads the saved settings. Never returns null and never throws:
        /// a missing or unreadable file yields default settings.
        /// </summary>
        public static AddInSettings Load()
        {
            try
            {
                lock (SyncRoot)
                {
                    if (!File.Exists(SettingsPath))
                        return new AddInSettings();

                    string json = File.ReadAllText(SettingsPath);
                    return JsonConvert.DeserializeObject<AddInSettings>(json) ?? new AddInSettings();
                }
            }
            catch
            {
                return new AddInSettings();
            }
        }

        /// <summary>
        /// Persists the settings, creating the settings folder when needed.
        /// </summary>
        public static void Save(AddInSettings settings)
        {
            if (settings == null)
                throw new ArgumentNullException(nameof(settings));

            lock (SyncRoot)
            {
                Directory.CreateDirectory(SettingsDirectory);
                File.WriteAllText(SettingsPath, JsonConvert.SerializeObject(settings, Formatting.Indented));
            }
        }
    }
}

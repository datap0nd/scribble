using System;
using Microsoft.Win32;

namespace Scribble.Configuration
{
    // Machine/user policy switches for Scribble, read from
    // Software\Policies\Scribble in HKLM (admin- or GPO-set, wins) and
    // HKCU. Policies only ever remove capabilities - nothing here
    // can widen what the add-in may do.
    public static class AdminPolicy
    {
        public const string PolicyKeyPath =
            "Software\\Policies\\Scribble";

        private static readonly string LegacyPolicyKeyPath =
            "Software\\Policies\\" + "AI" + "365";

        // Direct Gemini (Google browser sign-in and the Gemini
        // models in the picker) is an end-user capability when the
        // build defines SCRIBBLE_DIRECT_GEMINI, which Scribble.csproj
        // does for both configurations. Removing that define turns
        // the whole capability off again. The registry policy below
        // stays a one-way kill switch on top of the build gate.
        public static bool GeminiEnabledForEndUsers
        {
            get
            {
#if SCRIBBLE_DIRECT_GEMINI
                return true;
#else
                return false;
#endif
            }
        }

        // DisableGemini = 1 hides and blocks Google Gemini sign-in
        // across the suite; only the user's own OpenAI-compatible
        // endpoint remains available.
        public static bool GeminiDisabled
        {
            get
            {
                return !GeminiEnabledForEndUsers ||
                       ReadFlag("DisableGemini");
            }
        }

        private static bool ReadFlag(string valueName)
        {
            return ReadFlag(
                       Registry.LocalMachine,
                       PolicyKeyPath,
                       valueName) ||
                   ReadFlag(
                       Registry.CurrentUser,
                       PolicyKeyPath,
                       valueName) ||
                   ReadFlag(
                       Registry.LocalMachine,
                       LegacyPolicyKeyPath,
                       valueName) ||
                   ReadFlag(
                       Registry.CurrentUser,
                       LegacyPolicyKeyPath,
                       valueName);
        }

        private static bool ReadFlag(
            RegistryKey hive,
            string keyPath,
            string valueName)
        {
            try
            {
                using (var key = hive.OpenSubKey(keyPath))
                {
                    if (key == null)
                    {
                        return false;
                    }

                    var value = key.GetValue(valueName);
                    return value != null &&
                           Convert.ToInt32(value) == 1;
                }
            }
            catch
            {
                return false;
            }
        }
    }
}

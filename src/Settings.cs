using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using Microsoft.Win32;

namespace PptSelectionExport
{
    /// <summary>User choices, remembered between PowerPoint sessions (HKCU\Software\PptSelectionExport).</summary>
    internal sealed class Settings
    {
        private const string KeyPath = @"Software\PptSelectionExport";

        public bool Transparent = true;
        public Color Color = Color.White;
        public bool OpenFolder = true;
        public string LastDir = "";
        /// <summary>Resolution for PNG / JPG export, in pixels per inch.</summary>
        public int Dpi = 300;
        /// <summary>Whether the last picture was saved as JPEG (preselects the file type next time).</summary>
        public bool ImageJpeg = false;
        /// <summary>Align and distribute relative to the first selected object instead of the whole selection.</summary>
        public bool AlignToObject = false;
        /// <summary>Colours chosen in the colour picker, most recent first.</summary>
        public readonly List<Color> RecentColors = new List<Color>();
        private const int MaxRecent = 12;

        public void AddRecent(Color c)
        {
            c = Color.FromArgb(255, c);
            RecentColors.RemoveAll(x => x.ToArgb() == c.ToArgb());
            RecentColors.Insert(0, c);
            if (RecentColors.Count > MaxRecent) RecentColors.RemoveRange(MaxRecent, RecentColors.Count - MaxRecent);
        }

        public static Settings Load()
        {
            var s = new Settings();
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(KeyPath))
                {
                    if (k == null) return s;
                    s.Transparent = Convert.ToInt32(k.GetValue("Transparent", 1), CultureInfo.InvariantCulture) != 0;
                    s.OpenFolder = Convert.ToInt32(k.GetValue("OpenFolder", 1), CultureInfo.InvariantCulture) != 0;
                    s.LastDir = Convert.ToString(k.GetValue("LastDir", ""), CultureInfo.InvariantCulture);
                    s.ImageJpeg = Convert.ToInt32(k.GetValue("ImageJpeg", 0), CultureInfo.InvariantCulture) != 0;
                    s.AlignToObject = Convert.ToInt32(k.GetValue("AlignToObject", 0), CultureInfo.InvariantCulture) != 0;
                    foreach (string hex in Convert.ToString(k.GetValue("RecentColors", ""), CultureInfo.InvariantCulture).Split(','))
                    {
                        Color rc;
                        if (s.RecentColors.Count < MaxRecent && TryParseHex(hex, out rc)) s.RecentColors.Add(rc);
                    }
                    int dpi = Convert.ToInt32(k.GetValue("Dpi", 300), CultureInfo.InvariantCulture);
                    if (dpi >= 48 && dpi <= 1200) s.Dpi = dpi;
                    Color c;
                    if (TryParseHex(Convert.ToString(k.GetValue("Color", "FFFFFF"), CultureInfo.InvariantCulture), out c)) s.Color = c;
                }
            }
            catch { }
            return s;
        }

        public void Save()
        {
            try
            {
                using (RegistryKey k = Registry.CurrentUser.CreateSubKey(KeyPath))
                {
                    k.SetValue("Transparent", Transparent ? 1 : 0, RegistryValueKind.DWord);
                    k.SetValue("OpenFolder", OpenFolder ? 1 : 0, RegistryValueKind.DWord);
                    k.SetValue("LastDir", LastDir ?? "", RegistryValueKind.String);
                    k.SetValue("Dpi", Dpi, RegistryValueKind.DWord);
                    k.SetValue("ImageJpeg", ImageJpeg ? 1 : 0, RegistryValueKind.DWord);
                    k.SetValue("AlignToObject", AlignToObject ? 1 : 0, RegistryValueKind.DWord);
                    k.SetValue("RecentColors", string.Join(",", RecentColors.ConvertAll(ToHex).ToArray()), RegistryValueKind.String);
                    k.SetValue("Color", ToHex(Color), RegistryValueKind.String);
                }
            }
            catch { }
        }

        public static string ToHex(Color c)
        {
            return string.Format(CultureInfo.InvariantCulture, "{0:X2}{1:X2}{2:X2}", c.R, c.G, c.B);
        }

        public static bool TryParseHex(string hex, out Color c)
        {
            c = Color.White;
            if (hex == null) return false;
            hex = hex.Trim().TrimStart('#');
            int v;
            if (hex.Length != 6 || !int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out v)) return false;
            c = Color.FromArgb((v >> 16) & 0xFF, (v >> 8) & 0xFF, v & 0xFF);
            return true;
        }
    }
}

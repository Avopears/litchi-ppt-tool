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
        /// <summary>What the style buttons last applied; shown on the buttons. "None" = no fill / no outline.</summary>
        public Color FillColor = Color.FromArgb(0x3B, 0x82, 0xF6);
        public bool FillNone = false;
        public Color OutlineColor = Color.Black;
        public bool OutlineNone = false;
        public float OutlineWeight = 1.5f;
        public int OutlineDash = 0;
        /// <summary>Index into LineCaps.Caps; PowerPoint draws new lines with flat caps.</summary>
        public int OutlineCap = 2;
        /// <summary>"Slim down": pixels per inch the pictures keep (0 = all of them), JPEG quality, and whether photos may turn from PNG into JPEG.</summary>
        public int SlimPpi = 220;
        public int SlimQuality = 80;
        public bool SlimToJpeg = true;
        /// <summary>"Slide size": whether typing a width changes the height along with it.</summary>
        public bool SizeLockRatio = false;
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
                    Color sc;
                    if (TryParseHex(Convert.ToString(k.GetValue("FillColor", ""), CultureInfo.InvariantCulture), out sc)) s.FillColor = sc;
                    if (TryParseHex(Convert.ToString(k.GetValue("OutlineColor", ""), CultureInfo.InvariantCulture), out sc)) s.OutlineColor = sc;
                    s.FillNone = Convert.ToInt32(k.GetValue("FillNone", 0), CultureInfo.InvariantCulture) != 0;
                    s.OutlineNone = Convert.ToInt32(k.GetValue("OutlineNone", 0), CultureInfo.InvariantCulture) != 0;
                    float weight;
                    if (float.TryParse(Convert.ToString(k.GetValue("OutlineWeight", "1.5"), CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture, out weight)
                        && weight >= StyleTools.MinWeight && weight <= StyleTools.MaxWeight) s.OutlineWeight = weight;
                    int dash = Convert.ToInt32(k.GetValue("OutlineDash", 0), CultureInfo.InvariantCulture);
                    if (dash >= 0 && dash < StyleTools.Dashes.Length) s.OutlineDash = dash;
                    int cap = Convert.ToInt32(k.GetValue("OutlineCap", 2), CultureInfo.InvariantCulture);
                    if (cap >= 0 && cap < LineCaps.Caps.Length) s.OutlineCap = cap;
                    int slimPpi = Convert.ToInt32(k.GetValue("SlimPpi", 220), CultureInfo.InvariantCulture);
                    if (slimPpi == 0 || (slimPpi >= 72 && slimPpi <= 600)) s.SlimPpi = slimPpi;
                    int slimQuality = Convert.ToInt32(k.GetValue("SlimQuality", 80), CultureInfo.InvariantCulture);
                    if (slimQuality >= 30 && slimQuality <= 100) s.SlimQuality = slimQuality;
                    s.SlimToJpeg = Convert.ToInt32(k.GetValue("SlimToJpeg", 1), CultureInfo.InvariantCulture) != 0;
                    s.SizeLockRatio = Convert.ToInt32(k.GetValue("SizeLockRatio", 0), CultureInfo.InvariantCulture) != 0;
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
                    k.SetValue("FillColor", ToHex(FillColor), RegistryValueKind.String);
                    k.SetValue("FillNone", FillNone ? 1 : 0, RegistryValueKind.DWord);
                    k.SetValue("OutlineColor", ToHex(OutlineColor), RegistryValueKind.String);
                    k.SetValue("OutlineNone", OutlineNone ? 1 : 0, RegistryValueKind.DWord);
                    k.SetValue("OutlineWeight", OutlineWeight.ToString("0.##", CultureInfo.InvariantCulture), RegistryValueKind.String);
                    k.SetValue("OutlineDash", OutlineDash, RegistryValueKind.DWord);
                    k.SetValue("OutlineCap", OutlineCap, RegistryValueKind.DWord);
                    k.SetValue("SlimPpi", SlimPpi, RegistryValueKind.DWord);
                    k.SetValue("SlimQuality", SlimQuality, RegistryValueKind.DWord);
                    k.SetValue("SlimToJpeg", SlimToJpeg ? 1 : 0, RegistryValueKind.DWord);
                    k.SetValue("SizeLockRatio", SizeLockRatio ? 1 : 0, RegistryValueKind.DWord);
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

/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Utils;

namespace NanoXLSX.Extensions
{
    /// <summary>
    /// Represents color information for text formatting.
    /// </summary>
    public class Color
    {
        /// <summary>
        /// Enumeration for color schemes.
        /// </summary>
        public enum ColorScheme
        {
            /// <summary>Theme color that defines the dark1 (dk1) attribute of a theme.</summary>
            Dark1 = 0,
            /// <summary>Theme color that defines the light1 (lt1) attribute of a theme.</summary>
            Light1 = 1,
            /// <summary>Theme color that defines the dark2 (dk2) attribute of a theme.</summary>
            Dark2 = 2,
            /// <summary>Theme color that defines the light2 (lt2) attribute of a theme.</summary>
            Light2 = 3,
            /// <summary>Theme color that defines the accent1 attribute of a theme.</summary>
            Accent1 = 4,
            /// <summary>Theme color that defines the accent2 attribute of a theme.</summary>
            Accent2 = 5,
            /// <summary>Theme color that defines the accent3 attribute of a theme.</summary>
            Accent3 = 6,
            /// <summary>Theme color that defines the accent4 attribute of a theme.</summary>
            Accent4 = 7,
            /// <summary>Theme color that defines the accent5 attribute of a theme.</summary>
            Accent5 = 8,
            /// <summary>Theme color that defines the accent6 attribute of a theme.</summary>
            Accent6 = 9,
            /// <summary>Theme color that defines the hyperlink (hlink) attribute of a theme.</summary>
            Hyperlink = 10,
            /// <summary>Theme color that defines the followedHyperlink (folHlink) attribute of a theme.</summary>
            FollowedHyperlink = 11
        }

        /// <summary>
        /// Legacy OOXML / Excel indexed color palette.
        /// <para>
        /// This palette exists for backward compatibility with older Excel formats.
        /// Indices 0–7 are redundant with 8–15.
        /// </para>
        /// </summary>
        public enum IndexedColor : byte
        {
            /// <summary>Black (duplicate of index 8).</summary>
            Black0 = 0,
            /// <summary>White (duplicate of index 9).</summary>
            White1 = 1,
            /// <summary>Red (duplicate of index 10).</summary>
            Red2 = 2,
            /// <summary>Bright green (duplicate of index 11).</summary>
            BrightGreen3 = 3,
            /// <summary>Blue (duplicate of index 12).</summary>
            Blue4 = 4,
            /// <summary>Yellow (duplicate of index 13).</summary>
            Yellow5 = 5,
            /// <summary>Magenta (duplicate of index 14).</summary>
            Magenta6 = 6,
            /// <summary>Cyan (duplicate of index 15).</summary>
            Cyan7 = 7,
            /// <summary>Black (#000000).</summary>
            Black = 8,
            /// <summary>White (#FFFFFF).</summary>
            White = 9,
            /// <summary>Red (#FF0000).</summary>
            Red = 10,
            /// <summary>Bright green (#00FF00).</summary>
            BrightGreen = 11,
            /// <summary>Blue (#0000FF).</summary>
            Blue = 12,
            /// <summary>Yellow (#FFFF00).</summary>
            Yellow = 13,
            /// <summary>Magenta / Fuchsia (#FF00FF).</summary>
            Magenta = 14,
            /// <summary>Cyan / Aqua (#00FFFF).</summary>
            Cyan = 15,
            /// <summary>Dark red / maroon (#800000).</summary>
            DarkRed = 16,
            /// <summary>Dark green (#008000).</summary>
            DarkGreen = 17,
            /// <summary>Dark blue / navy (#000080).</summary>
            DarkBlue = 18,
            /// <summary>Olive (#808000).</summary>
            Olive = 19,
            /// <summary>Purple (#800080).</summary>
            Purple = 20,
            /// <summary>Teal (#008080).</summary>
            Teal = 21,
            /// <summary>Light gray / silver (#C0C0C0).</summary>
            LightGray = 22,
            /// <summary>Medium gray (#808080).</summary>
            Gray = 23,
            /// <summary>Light cornflower blue (#9999FF).</summary>
            LightCornflowerBlue = 24,
            /// <summary>Dark rose (#993366).</summary>
            DarkRose = 25,
            /// <summary>Light yellow (#FFFFCC).</summary>
            LightYellow = 26,
            /// <summary>Light cyan (#CCFFFF).</summary>
            LightCyan = 27,
            /// <summary>Dark purple (#660066).</summary>
            DarkPurple = 28,
            /// <summary>Salmon pink (#FF8080).</summary>
            Salmon = 29,
            /// <summary>Medium blue (#0066CC).</summary>
            MediumBlue = 30,
            /// <summary>Light lavender blue (#CCCCFF).</summary>
            LightLavender = 31,
            /// <summary>Dark navy blue (#000080).</summary>
            Navy = 32,
            /// <summary>Strong magenta (#FF00FF).</summary>
            StrongMagenta = 33,
            /// <summary>Strong yellow (#FFFF00).</summary>
            StrongYellow = 34,
            /// <summary>Strong cyan (#00FFFF).</summary>
            StrongCyan = 35,
            /// <summary>Dark violet (#800080).</summary>
            DarkViolet = 36,
            /// <summary>Dark maroon (#800000).</summary>
            DarkMaroon = 37,
            /// <summary>Dark teal (#008080).</summary>
            DarkTeal = 38,
            /// <summary>Pure blue (#0000FF).</summary>
            PureBlue = 39,
            /// <summary>Sky blue (#00CCFF).</summary>
            SkyBlue = 40,
            /// <summary>Pale cyan (#CCFFFF).</summary>
            PaleCyan = 41,
            /// <summary>Light mint green (#CCFFCC).</summary>
            LightMint = 42,
            /// <summary>Light pastel yellow (#FFFF99).</summary>
            PastelYellow = 43,
            /// <summary>Light sky blue (#99CCFF).</summary>
            LightSkyBlue = 44,
            /// <summary>Rose pink (#FF99CC).</summary>
            Rose = 45,
            /// <summary>Lavender (#CC99FF).</summary>
            Lavender = 46,
            /// <summary>Peach (#FFCC99).</summary>
            Peach = 47,
            /// <summary>Royal blue (#3366FF).</summary>
            RoyalBlue = 48,
            /// <summary>Turquoise (#33CCCC).</summary>
            Turquoise = 49,
            /// <summary>Light olive green (#99CC00).</summary>
            LightOlive = 50,
            /// <summary>Gold (#FFCC00).</summary>
            Gold = 51,
            /// <summary>Orange (#FF9900).</summary>
            Orange = 52,
            /// <summary>Dark orange (#FF6600).</summary>
            DarkOrange = 53,
            /// <summary>Blue gray (#666699).</summary>
            BlueGray = 54,
            /// <summary>Medium gray (#969696).</summary>
            MediumGray = 55,
            /// <summary>Dark slate blue (#003366).</summary>
            DarkSlateBlue = 56,
            /// <summary>Sea green (#339966).</summary>
            SeaGreen = 57,
            /// <summary>Very dark green (#003300).</summary>
            VeryDarkGreen = 58,
            /// <summary>Dark olive (#333300).</summary>
            DarkOlive = 59,
            /// <summary>Brown (#993300).</summary>
            Brown = 60,
            /// <summary>Dark rose (duplicate of index 25).</summary>
            DarkRoseDuplicate = 61,
            /// <summary>Indigo / dark blue-purple (#333399).</summary>
            Indigo = 62,
            /// <summary>Very dark gray (#333333).</summary>
            VeryDarkGray = 63,
            /// <summary>
            /// System foreground color.
            /// <para>
            /// The actual color is determined by the host operating system or theme.
            /// </para>
            /// </summary>
            SystemForeground = 64,
            /// <summary>
            /// System background color.
            /// <para>
            /// The actual color is determined by the host operating system or theme.
            /// </para>
            /// </summary>
            SystemBackground = 65
        }

        private string rgb;

        /// <summary>
        /// Indicating whether the color is automatic and system color dependent.
        /// </summary>
        public bool? Auto { get; set; }
        /// <summary>
        /// Indexed color value. Only used for backwards compatibility.
        /// </summary>
        public IndexedColor? Indexed { get; set; }
        /// <summary>
        /// Gets or sets the ARGB color value as a string. If the leading alpha value is omitted (6 characters), it is set to 'FF' (opaque) automatically.
        /// </summary>
        public string Rgb
        {
            get { return rgb; }
            set
            {
                if (value != null && value.Length == 6)
                {
                    value = "FF" + value;
                }
                Validators.ValidateColor(value, true, true);
                rgb = value;
            }
        }
        /// <summary>
        /// Gets or sets the theme identifier for the user interface.
        /// </summary>
        public ColorScheme? Theme { get; set; }
        /// <summary>
        /// Gets or sets the tint value for the color, where -1.0 is darkest and 1.0 is lightest (0.0 = normal).
        /// </summary>

#pragma warning disable CA1805 // Do not initialize unnecessarily
        public double Tint { get; set; } = 0.0;
#pragma warning restore CA1805 // Do not initialize unnecessarily

        /// <summary>
        /// Creates a Color instance from an RGB hex string. The leading alpha value is set to 'AA' (opaque) automatically.
        /// </summary>
        /// <param name="rgb">The RGB hex string (e.g. '#FF0000' or 'FF0000' for red)</param>
        /// <returns>Color instance</returns>
        public static Color FromRgb(string rgb)
        {
            rgb = "AA" + rgb.TrimStart('#');
            Validators.ValidateColor(rgb, false);
            return new Color { Rgb = rgb };
        }

        /// <summary>
        /// Creates a Color instance from an ARGB hex string.
        /// </summary>
        /// <param name="argb">The ARGB hex string (e.g. '#FFFF0000' or 'FFFF0000' for red)</param>
        /// <returns>Color instance</returns>
        public static Color FromArgb(string argb)
        {
            argb = argb.TrimStart('#');
            Validators.ValidateColor(argb, true);
            return new Color { Rgb = argb };
        }

        /// <summary>
        /// Creates a Color instance from a theme color with an optional tint.
        /// </summary>
        /// <param name="theme">The theme color</param>
        /// <param name="tint">The tint value, where -1.0 is darkest and 1.0 is lightest (0.0 = normal)</param>
        /// <returns>Color instance</returns>
        public static Color FromTheme(ColorScheme theme, double tint = 0.0)
        {
            return new Color { Theme = theme, Tint = tint };
        }

        /// <summary>
        /// Creates a Color instance from an indexed color value.
        /// </summary>
        /// <param name="indexed">The indexed color value</param>
        /// <returns>Color instance</returns>
        public static Color FromIndexed(IndexedColor indexed)
        {
            return new Color { Indexed = indexed };
        }

        /// <summary>
        /// Creates a Color instance representing automatic color.
        /// </summary>
        /// <returns>Color instance</returns>
        public static Color Automatic()
        {
            return new Color { Auto = true };
        }

        /// <summary>
        /// Creates a copy of the current Color instance.
        /// </summary>
        /// <returns>Color instance</returns>
        public Color Copy()
        {
            return new Color
            {
                Auto = this.Auto,
                Indexed = this.Indexed,
                Rgb = this.Rgb,
                Theme = this.Theme,
                Tint = this.Tint
            };
        }

    }
}

using System;
using System.Collections.Generic;
using System.Text;

namespace NanoXLSX.Extensions
{
    /// <summary>
    /// Represents color information for text formatting.
    /// </summary>
    public class Color
    {
        public bool? Auto { get; set; }
        public uint? Indexed { get; set; }
        public string Rgb { get; set; }
        public uint? Theme { get; set; }
        public double Tint { get; set; } = 0.0;

        public static Color FromRgb(string rgb)
        {
            return new Color { Rgb = rgb };
        }

        public static Color FromTheme(uint theme, double tint = 0.0)
        {
            return new Color { Theme = theme, Tint = tint };
        }

        public static Color FromIndexed(uint indexed)
        {
            return new Color { Indexed = indexed };
        }

        public static Color Automatic()
        {
            return new Color { Auto = true };
        }

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

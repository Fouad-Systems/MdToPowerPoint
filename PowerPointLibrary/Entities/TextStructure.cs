using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.Entities
{
    /// <summary>
    /// a Texte part or element , description
    /// </summary>
    public class TextElementStyle
    {
        public enum TextStyles
        {
            Blod,
            Italic
        }
        public int Start;
        public int Length;
        public TextStyles TextStyle;

        public TextElementStyle(int Start, int Length, TextStyles TextStyle)
        {
            this.Start = Start;
            this.Length = Length;
            this.TextStyle = TextStyle;
        }

    }

    public class TextStructure
    {
        public string Text { get; set; }


        public List<TextElementStyle> TextElementStyles = new List<TextElementStyle>();

    }
}

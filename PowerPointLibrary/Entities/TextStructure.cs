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
    public class TextElementStyle : ICloneable
    {
        public enum TextStyles
        {
            Blod,
            Italic
        }
        public int Start;
        public int Length;
        public TextStyles TextStyle;

        public TextElementStyle()
        {

        }
        public TextElementStyle(int Start, int Length, TextStyles TextStyle)
        {
            this.Start = Start;
            this.Length = Length;
            this.TextStyle = TextStyle;
        }

        public object Clone()
        {
            TextElementStyle clone = new TextElementStyle();
            clone.Start = this.Start;
            clone.Length = this.Length;
            clone.TextStyle = TextStyle;
            return clone;

        }
    }

    public class TextStructure : ICloneable
    {
        public override string ToString()
        {
            return this.Text;
        }
        public string Text { get; set; }


        public List<TextElementStyle> TextElementStyles = new List<TextElementStyle>();

        public object Clone()
        {
            TextStructure clone = new TextStructure();
            clone.Text = Text;
            clone.TextElementStyles = TextElementStyles.Select(o => o.Clone() as TextElementStyle).ToList() ;
            return clone;
        }
    }
}

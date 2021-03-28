using Microsoft.Toolkit.Parsers.Markdown;
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

        public bool IsBlod { get; set; }
        public bool IsItalic { get; set; }
        public bool IsBullet { get; set; }

   

        public string FontColor { get; set; }

        public int Start;
        public int Length;

        public ListStyle ListStyle;

      
        public TextElementStyle(int Start, int Length)
        {
            this.Start = Start;
            this.Length = Length;
        }

        public object Clone()
        {
            TextElementStyle clone = new TextElementStyle(this.Start,this.Length);
            clone.Start = this.Start;
            clone.Length = this.Length;
            clone.IsBlod = IsBlod;
            clone.IsItalic = IsItalic;
            clone.IsBullet = IsBullet;
            clone.FontColor = FontColor;
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


        public TextStructure()
        {
            Text = "";
        }

        public object Clone()
        {
            TextStructure clone = new TextStructure();
            clone.Text = Text;
            clone.TextElementStyles = TextElementStyles.Select(o => o.Clone() as TextElementStyle).ToList() ;
            return clone;
        }
    }
}

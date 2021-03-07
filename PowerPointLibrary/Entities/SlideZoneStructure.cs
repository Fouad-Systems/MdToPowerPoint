using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.Entities
{
    public class SlideZoneStructure : ICloneable
    {
        public string Name { get; set; }

        public TextStructure Text { get; set; }

        public object Clone()
        {
            SlideZoneStructure clone = new SlideZoneStructure();
            clone.Name = Name;
            if (Text != null)
                clone.Text = Text.Clone() as TextStructure;
            return clone;

        }

        // public ImageStructure { get; set; }
    }
}

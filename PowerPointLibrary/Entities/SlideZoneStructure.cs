using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.Entities
{
    public class SlideZoneStructure : ICloneable
    {
        public override string ToString()
        {
            return this.Name;
        }
        public string Name { get; set; }

        public TextStructure Text { get; set; }

        public List<ImageStructure> Images { get; set; }

        public object Clone()
        {
            SlideZoneStructure clone = new SlideZoneStructure();
            clone.Name = Name;
            if (Text != null)
                clone.Text = Text.Clone() as TextStructure;

            if (Images != null)
                clone.Images = Images.Select(m => m.Clone() as ImageStructure).ToList();
            return clone;

        }

        // public ImageStructure { get; set; }
    }
}

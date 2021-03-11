using PowerPointLibrary.Entities.Enums;
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

        public int Order { set; get; }

        public List<ContentTypes> ContentTypes { get; set; }


        public TextStructure Text { get; set; }

        public List<ImageStructure> Images { get; set; }


        public SlideZoneStructure()
        {
            ContentTypes = new List<ContentTypes>();
        }



        public object Clone()
        {
            SlideZoneStructure clone = new SlideZoneStructure();
            clone.Name = Name;
            clone.Order = Order;
            if (Text != null)
                clone.Text = Text.Clone() as TextStructure;

            if (Images != null)
                clone.Images = Images.Select(m => m.Clone() as ImageStructure).ToList();

            clone.ContentTypes = ContentTypes.Select(o => o).ToList(); ;

            return clone;

        }

        // public ImageStructure { get; set; }
    }
}

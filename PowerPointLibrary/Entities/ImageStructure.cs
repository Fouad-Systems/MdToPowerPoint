using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.Entities
{
    public class ImageStructure : ICloneable
    {

        public string Url { get; set; }
      

        public string Tooltip { get; set; }

        public int ImageWidth { get; set; }

        public int ImageHeight { get; set; }

        public object Clone()
        {
            ImageStructure clone = new ImageStructure();
            clone.Url = this.Url;
            clone.Tooltip = this.Tooltip;
            clone.ImageHeight = ImageHeight;
            clone.ImageHeight = ImageHeight;
            return clone;
        }
    }
}

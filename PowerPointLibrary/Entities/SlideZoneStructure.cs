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


        /// <summary>
        /// Line number in generated layout
        /// </summary>
        public int GLines { get; set; }

        /// <summary>
        /// Columns number in generated layout
        /// </summary>
        public int GColumns { get; set; }

        public List<ContentTypes> ContentTypes { get; set; }


        public TextStructure Text { get; set; }

        public List<ImageStructure> Images { get; set; }


        // Postion in layout
        public float Width { get; internal set; }

        // Postion in layout
        public float Height { get; internal set; }

        // Postion in layout
        public float Top { get; internal set; }

        // Postion in layout
        public float Left { get; internal set; }
        public int Row { get; internal set; }

        public SlideZoneStructure()
        {
            ContentTypes = new List<ContentTypes>();
        }


        public void Clone(SlideZoneStructure clone)
        {
            clone.Name = Name;
            clone.Order = Order;
            clone.Width = Width;
            clone.Height = Height;
            clone.Left = Left;
            clone.Top = Top;


            if (Text != null)
                clone.Text = Text.Clone() as TextStructure;

            if (Images != null)
                clone.Images = Images.Select(m => m.Clone() as ImageStructure).ToList();

            clone.ContentTypes = ContentTypes.Select(o => o).ToList(); ;
        }
        public object Clone()
        {
            SlideZoneStructure clone = new SlideZoneStructure();
            this.Clone(clone);

            return clone;

        }

        public bool IsEmpty()
        {
            if (
                ( this.Text != null && !string.IsNullOrEmpty(this.Text.Text))
                 || (this.Images != null && this.Images.Count() >= 1 )
               )
                return false;

            return true;
        }

        public bool IsImage()
        {
            if (this.Images != null && this.Images.Count() >= 1)
                return true;
            else
                return false;
        }

        public bool IsTitle()
        {
            if (this.ContentTypes.Contains(Entities.Enums.ContentTypes.Title))
                return true;
            else
                return false;
          
        }

      

        // public ImageStructure { get; set; }
    }
}

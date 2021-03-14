using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.Entities
{



    public class PresentationStructure
    {

        public PresentationStructure _TemplateStructure { get; set; }

        public List<SlideStructure> Slides { set; get; }


        public PresentationStructure()
        {
            this.Slides = new List<SlideStructure>();
        }


        public SlideStructure CurrentSlide
        {
            get
            {
                return this.Slides.Last();
            }
        }

    }
}

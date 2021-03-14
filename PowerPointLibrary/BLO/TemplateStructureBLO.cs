using PowerPointLibrary.Entities;
using PowerPointLibrary.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.BLO
{
    public class TemplateStructureBLO
    {
        private PresentationStructure _PresentationStructure;

        public TemplateStructureBLO(PresentationStructure presentationStructure)
        {
            _PresentationStructure = presentationStructure;
        }

        public SlideStructure GetSlide(string layout)
        {
            var slide = this._PresentationStructure._TemplateStructure.Slides.Where(s => s.Layout == layout)
                .FirstOrDefault();
            if (slide == null) throw new PplException($"The layout '{layout}' dones't exist");
            return slide;
        }
    }
}

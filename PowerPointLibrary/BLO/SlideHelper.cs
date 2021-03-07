using Microsoft.Office.Interop.PowerPoint;
using PowerPointLibrary.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.BLO
{

   

    //
    // default layout : Titre de section, Title Slide,Two Content,Console 1,Console 2, Console-Répertoire 1
    //

    public class SlideHelper
    {
        public PresentationBLO presentationHelper;
        public Slide Slide;

        public SlideHelper(PresentationBLO presentationHelper, int index)
        {
            this.presentationHelper = presentationHelper;
            Slide = presentationHelper._Presentation.Slides.Add(index, PpSlideLayout.ppLayoutBlank);
            Slide.CustomLayout = presentationHelper._Presentation.Designs[1].SlideMaster.CustomLayouts[1];
          
        }

        public void ChangeLayout(string LayoutName)
        {
            CustomLayout customLayout = this.FindCustomLayoutByName(LayoutName);
            if (customLayout == null) throw new PowerPointLibraryException($"The layout {LayoutName} doesn't exist");
            Slide.CustomLayout = customLayout;
        }

        public CustomLayout FindCustomLayoutByName(string Name)
        {
            foreach (CustomLayout customLayout in this.Slide.Design.SlideMaster.CustomLayouts)
            {
                if (customLayout.Name == Name) return customLayout;
            }
            return null;
        }



    }
}

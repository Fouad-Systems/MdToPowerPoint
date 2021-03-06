using Microsoft.Office.Interop.PowerPoint;
using PowerPointLibrary.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.Helpers
{

   

    //
    // default layout : Titre de section, Title Slide,Two Content,Console 1,Console 2, Console-Répertoire 1
    //

    public class SlideHelper
    {
        public PresentationHelper presentationHelper;
        public Slide Slide;

        public SlideHelper(PresentationHelper presentationHelper, int index)
        {
            this.presentationHelper = presentationHelper;
            Slide = presentationHelper.oPresentation.Slides.Add(index, PpSlideLayout.ppLayoutBlank);
            Slide.CustomLayout = presentationHelper.oPresentation.Designs[1].SlideMaster.CustomLayouts[1];
          
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

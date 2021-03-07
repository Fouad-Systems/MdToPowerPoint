using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.BLO
{
    public class TitleAndContentSlideHelper : SlideHelper
    {
        public string CurrentShapesName { get; set; }

        public TitleAndContentSlideHelper(PresentationBLO presentationHelper, int index):base(presentationHelper, index)
        {
            //ChangeLayout("Titre et contenu");
            ChangeLayout("Console 2");
        }

        public String Title {
            set
            {
                TextRange oText = this.Slide.Shapes[1].TextFrame.TextRange;
                // oText.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                oText.Text = value;
            }
        
        }
        public String Content {
            set
            {
                TextRange oText = this.Slide.Shapes[2].TextFrame.TextRange;
                // oText.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                oText.Text = value;


            }
            get
            {
                TextRange oText = this.Slide.Shapes[2].TextFrame.TextRange;
                // oText.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                return oText.Text;
            }
        }

       
    }
}

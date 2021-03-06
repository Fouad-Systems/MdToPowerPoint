using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.Helpers
{
    public class PresentationHelper
    {
        internal Application oPowerPoint;
        internal Presentation oPresentation;

        public PresentationHelper()
        {
            oPowerPoint = new Application();
            // By default PowerPoint is invisible, till you make it visible: 
            // oPowerPoint.Visible = MsoTriState.msoCTrue;
        }

        /// <summary>
        /// Create a new Presentation. 
        /// </summary>
        /// <param name="TemplateName"></param>
        public void Create(string TemplateName)
        {


            oPresentation = oPowerPoint.Presentations.Add(MsoTriState.msoFalse);
            oPresentation.ApplyTemplate(TemplateName);
        }

        public void SaveAs(string fileName)
        {
            this.oPresentation.SaveAs(fileName,
                    PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                    MsoTriState.msoTriStateMixed);

          

        }

        public void Close()
        {
            this.oPresentation.Close();
            this.oPowerPoint.Quit();
        }

    }
}

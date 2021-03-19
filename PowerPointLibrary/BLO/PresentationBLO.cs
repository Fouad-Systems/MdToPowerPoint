using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Toolkit.Parsers.Markdown;
using PowerPointLibrary.Helper;
using PowerPointLibrary.Manager;
using PowerPointLibrary.Entities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Toolkit.Parsers.Markdown.Blocks;
using PowerPointLibrary.Helper.Enumerations;
using System.Drawing;
using PowerPointLibrary.Exceptions;
using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;

namespace PowerPointLibrary.BLO
{
    /// <summary>
    /// Create presentation
    /// </summary>
    public class PresentationBLO
    {

        #region Attributes
        internal Microsoft.Office.Interop.PowerPoint.Application _Application;
        internal Presentation _Presentation;
        // internal PresentationStructure _TemplateStructure;
        internal PresentationStructure _PresentationStructure;

        private readonly PowerPointApplicationManager _ApplicationManager;
        private readonly PresentationManager _PresentationManager;
        private readonly SlideManager _SlideManager;
        private readonly ShapesManager _ShapeManager;
        private readonly TextRangeManager _TextRangeManager;

        private readonly PresentationStructureBLO _PresentationStructureBLO;
        private TemplateStructureBLO _TemplateStructureBLO;
        private readonly TextStructureBLO _TextStructureBLO;
        private readonly CommentActionBLO _CommentActionBLO;
        private readonly SlideBLO _SlideBLO;
        private readonly SlideZoneStructureBLO _SlideZoneStructureBLO;
        private GLayoutStructureBLO _GLayoutStructureBLO;

        GeneratePresentationBLO _GeneratePresentationBLO;

        public PplArguments pplArguments;
        #endregion

        public PresentationBLO(PplArguments pplArguments)
        {
            this.pplArguments = pplArguments;

            // Init Manager
            _ApplicationManager = new PowerPointApplicationManager();
            _PresentationManager = new PresentationManager();
            _SlideManager = new SlideManager();
            _TextRangeManager = new TextRangeManager();
            _ShapeManager = new ShapesManager();




            _Application = _ApplicationManager.CreatePowerPointApplication();
            _PresentationStructure = new PresentationStructure();
            _GLayoutStructureBLO = new GLayoutStructureBLO();

            // Init BLO
            _PresentationStructureBLO = new PresentationStructureBLO(this._PresentationStructure);
            _TextStructureBLO = new TextStructureBLO();
            _CommentActionBLO = new CommentActionBLO();
            _SlideBLO = new SlideBLO(_PresentationStructure);
            _SlideZoneStructureBLO = new SlideZoneStructureBLO();
            _TemplateStructureBLO = new TemplateStructureBLO(_PresentationStructure);

          
        }

        public void CreatePresentation(MarkdownDocument mdDocument)
        {

            _PresentationStructureBLO.CreatePresentationDataStructure(mdDocument);


            _GeneratePresentationBLO = new GeneratePresentationBLO(_Application, _Presentation, _PresentationStructure, pplArguments);


            _GeneratePresentationBLO.GeneratePresentation();

            this.SaveAs(this.pplArguments.OutPutFile);
            this.Close();

     
            
        }


        #region Create,Save,Close
        /// <summary>
        /// Create a new Presentation from template
        /// </summary>
        /// <param name="TemplateName"></param>
        public void Create()
        {
            //string CurrentDirectory = Environment.CurrentDirectory;
            //string PowerPointTemplateFileName = CurrentDirectory + "/" + TemplateName + ".pptx";
            //if (!File.Exists(PowerPointTemplateFileName))
            //{
            //    PowerPointTemplateFileName = CurrentDirectory + "/" + TemplateName + ".potx";

            //    if (!File.Exists(PowerPointTemplateFileName))
            //    {
            //        string msg = $"The file { PowerPointTemplateFileName} or {TemplateName + ".potx"} not exist";
            //        throw new PowerPointLibrary.Exceptions.PplException(msg);
            //    }
            //}

            // Open an existing PowerPoint presentation
            _Presentation = _PresentationManager.OpenExistingPowerPointPresentation(
                    _Application,
                    this.pplArguments.TemplatePath);


            this._PresentationStructure._TemplateStructure = _PresentationStructureBLO.LoadConfiguration(this.pplArguments.TemplateConfigurationPath);

        }

        public void SaveAs(string fileName)
        {
            this._Presentation.SaveAs(fileName,
                    PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                    MsoTriState.msoTriStateMixed);



        }

        public void Close()
        {
            this._Presentation.Close();
            this._Application.Quit();
        }

        #endregion


     


        private void PandocCommande(string v)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = @"powershell.exe";
            startInfo.Arguments = @" pandoc E:\formations\11.React\src\hello-world-react.md -o E:\formations\11.React\src\bb.pptx ";
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.UseShellExecute = false;
            startInfo.CreateNoWindow = true;
            Process process = new Process();
            process.StartInfo = startInfo;
            process.Start();
        }

        public SlideStructure CurrentSlide
        {
            get
            {
                return this._PresentationStructure.CurrentSlide;
            }
        }
    }
}

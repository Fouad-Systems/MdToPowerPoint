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

namespace PowerPointLibrary.BLO
{
    /// <summary>
    /// Create presentation
    /// </summary>
    public class PresentationBLO
    {

        internal Application _Application;
        internal Presentation _Presentation;
        internal PresentationStructure _TemplateStructure;
        internal PresentationStructure _PresentationStructure;

        #region Manager
        private readonly PowerPointApplicationManager _ApplicationManager = new PowerPointApplicationManager();
        private readonly PresentationManager _PresentationManager = new PresentationManager();

        private readonly SlideManager _SlideManager = new SlideManager();

        private readonly ShapesManager _ShapeManager = new ShapesManager();
        private readonly TextRangeManager _TextRangeManager = new TextRangeManager();

        #endregion

        private readonly PresentationStructureBLO _TemplateStructureBLO = new PresentationStructureBLO();
        private readonly TextStructureBLO _TextStructureBLO = new TextStructureBLO();

        public PresentationBLO()
        {
            _Application = _ApplicationManager.CreatePowerPointApplication();
            _PresentationStructure = new PresentationStructure();
        }

        /// <summary>
        /// Create a new Presentation from template
        /// </summary>
        /// <param name="TemplateName"></param>
        public void Create(string TemplateName)
        {
            string CurrentDirectory = Environment.CurrentDirectory;
            string PowerPointTemplateFileName = CurrentDirectory + "/" + TemplateName + ".pptx";
            if (!File.Exists(PowerPointTemplateFileName))
            {
                PowerPointTemplateFileName = CurrentDirectory + "/" + TemplateName + ".potx";

                if (!File.Exists(PowerPointTemplateFileName))
                {
                    string msg = $"The file { PowerPointTemplateFileName} or {TemplateName + ".potx"} not exist";
                    throw new PowerPointLibrary.Exceptions.PowerPointLibraryException(msg);
                }
            }

            // Open an existing PowerPoint presentation
            _Presentation = _PresentationManager.OpenExistingPowerPointPresentation(
                    _Application,
                    PowerPointTemplateFileName);


            this._TemplateStructure = _TemplateStructureBLO.LoadConfiguration(TemplateName);

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


        public void CreatePresentationDataStructure(MarkdownDocument mdDocument)
        {

            foreach (var element in mdDocument.Blocks)
            {
                if (element is HeaderBlock header)
                {

                    this.AddSlide(header);

                    var currentSlide = this._PresentationStructure.CurrentSlide;
                    SlideZoneStructure slideZoneTitle = currentSlide.SlideZones.Where(z => z.Name == "Titre")
                        .FirstOrDefault();
                    slideZoneTitle.Text = _TextStructureBLO.CreateFromMarkdownBlock(header);

                    //slide = new TitleAndContentSlideHelper(presentationBLO, SlideIndex++);
                    //string Slide_name = slide.Slide.Name;
                    //int c = slide.Slide.Shapes.Count;
                    //string name = slide.Slide.Shapes[1].Name;


                    //TextRange TitleTextRange = slide.Slide.Shapes[1].TextFrame.TextRange;
                    //new TextRangeHelper(TitleTextRange).AddMarkdownBlock(element);

                    //oText.Text = "Bonjour l'informatique";
                    //oText.Words(1, 1).Find("Bonjour").Font.Bold = MsoTriState.msoCTrue;
                    //oText.Words(1, 1).Find("Bonjour").Font.Size = 20;
                    //oText =  oText.Words(1,1).InsertAfter(oText.Words(1, 1));

                }

                if (element is ParagraphBlock Paragraph)
                {
                    //if (Paragraph.Inlines[0].Type == MarkdownInlineType.Comment)
                    //{
                    //    string comment = Paragraph.Inlines[0].ToString();

                    //    // Change Slide layout
                    //    if (comment.StartsWith("<!-- slide : "))
                    //    {
                    //        string layout = comment.Replace("<!-- slide : ", "");
                    //        layout = layout.Replace("-->", "");
                    //        layout = layout.Trim();
                    //        // slide = new TitleAndContentSlideHelper(presentationHelper, SlideIndex++);
                    //        slide.ChangeLayout(layout);
                    //    }

                    //    // Change zone
                    //    if (comment.StartsWith("<!-- zone : "))
                    //    {
                    //        string ShapesName = comment.Replace("<!-- zone : ", "");
                    //        ShapesName = ShapesName.Replace("-->", "");
                    //        ShapesName = ShapesName.Trim();
                    //        // slide = new TitleAndContentSlideHelper(presentationHelper, SlideIndex++);
                    //        slide.CurrentShapesName = ShapesName;
                    //    }

                    //}

                    //TextRange TitleTextRange = slide.Slide.Shapes[2].TextFrame.TextRange;
                    //if (!string.IsNullOrEmpty(slide.CurrentShapesName))
                    //{
                    //    string Slide_name = slide.Slide.Name;
                    //    int c = slide.Slide.Shapes.Count;
                    //    string name = slide.Slide.Shapes[1].Name;
                    //    name = slide.Slide.Shapes[2].Name;
                    //    name = slide.Slide.Shapes[3].Name;
                    //    TitleTextRange = slide.Slide.Shapes["Content Placeholder 6"].TextFrame.TextRange;
                    //}

                    var currentSlide = this._PresentationStructure.CurrentSlide;
                    SlideZoneStructure ContenuZone = currentSlide.SlideZones.Where(z => z.Name == "Contenu")
                        .FirstOrDefault();
                    if (ContenuZone != null)
                        ContenuZone.Text = _TextStructureBLO.CreateFromMarkdownBlock(Paragraph);
                }

            }
        }

        private void AddSlide(HeaderBlock header)
        {
            SlideStructure slideStructure = new SlideStructure();
            _PresentationStructure.Slides.Add(slideStructure);
            slideStructure.Name = "Slide" + this._PresentationStructure.Slides.Count;

            if (header.HeaderLevel == 1) slideStructure.Template = "Titre partie";
            if (header.HeaderLevel >= 2) slideStructure.Template = "Titre et contenue";

            // Add Template Zone to Slide
            var TemplateSlide = _TemplateStructure.Slides
                 .Where(s => s.Name == slideStructure.Template).FirstOrDefault();

            slideStructure.SlideZones = TemplateSlide.SlideZones.Select(s => new SlideZoneStructure() { Name = s.Name  }).ToList();
            slideStructure.TemplateSlide = TemplateSlide;
        }

        public void GeneratePresentation()
        {

            // Add Slides
            foreach (var slide in _PresentationStructure.Slides)
            {
                SlideRange slideRange = _SlideManager
                    .CloneSlide(_Presentation, _Presentation.Slides[slide.TemplateSlide.Order], Locations.Location.Last);

                foreach (var SlideZone in slide.SlideZones)
                {
                    // Add Text
                    if (SlideZone.Text != null)
                    {
                        var shape = slideRange.Shapes[SlideZone.Name];
                        _TextRangeManager.AddTextStructure(shape.TextFrame.TextRange, SlideZone.Text);
                    }
                }
            }

            // Delete Template Slide
          



            this.SaveAs(Environment.CurrentDirectory + "/" + "output.pptx");
            this.Close();

        }
    }
}

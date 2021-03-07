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
        private readonly PowerPointApplicationManager _ApplicationManager;
        private readonly PresentationManager _PresentationManager;
        private readonly SlideManager _SlideManager;
        private readonly ShapesManager _ShapeManager;
        private readonly TextRangeManager _TextRangeManager;

        #endregion

        private readonly PresentationStructureBLO _TemplateStructureBLO;
        private readonly TextStructureBLO _TextStructureBLO;
        private readonly CommentActionBLO _CommentActionBLO;
        private readonly SlideBLO _SlideBLO;

        public PresentationBLO()
        {
            // Init Manager
            _ApplicationManager = new PowerPointApplicationManager();
            _PresentationManager = new PresentationManager();
            _SlideManager = new SlideManager();
            _TextRangeManager = new TextRangeManager();


            // Init BLO
            _TemplateStructureBLO = new PresentationStructureBLO();
            _TextStructureBLO = new TextStructureBLO();
            _CommentActionBLO = new CommentActionBLO();
            _SlideBLO = new SlideBLO(this);


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
                    throw new PowerPointLibrary.Exceptions.PplException(msg);
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
                    string layout = "";
                    if (header.HeaderLevel == 1) layout = "Titre partie";
                    if (header.HeaderLevel >= 2) layout = "Titre et contenu";

                    _SlideBLO.AddSlide(layout);

                    SlideZoneStructure zoneTitle = this.CurrentSlide.CurrentZone;
                
                    if (zoneTitle != null)
                        zoneTitle.Text = _TextStructureBLO.CreateFromMarkdownBlock(header);

                }

                if (element is ParagraphBlock Paragraph)
                {
                 
                    _SlideBLO.ChangeZoneToParagraphe();

                    // if paragraphe is action
                    if (Paragraph.Inlines[0].Type == MarkdownInlineType.Comment
                        && _CommentActionBLO.IsAction(Paragraph.Inlines[0].ToString()))
                    {

                        string comment = Paragraph.Inlines[0].ToString();

                        CommentAction commentAction = _CommentActionBLO.ParseComment(comment);

                        switch (commentAction.ActionType)
                        {
                            case CommentAction.ActionTypes.ChangeLayout:
                                _SlideBLO
                                    .ChangeLayout(this.CurrentSlide, commentAction.Layout);
                                break;
                            case CommentAction.ActionTypes.ChangeZone:
                                _SlideBLO
                                    .ChangeCurrentZone(this.CurrentSlide, commentAction.ZoneName);
                                break;
                            case CommentAction.ActionTypes.NewSlide:
                                _SlideBLO.AddSlide(commentAction.Layout);
                                break;
                        }
                    }
                    else
                    {
                        
                        if (this.CurrentSlide.CurrentZone != null)
                        {
                            this.CurrentSlide.CurrentZone.Text = _TextStructureBLO.CreateFromMarkdownBlock(Paragraph);
                        }
                        else
                        {
                            if (this.CurrentSlide.AddToNotes)
                            {
                                this.CurrentSlide.Notes.Text = _TextStructureBLO.CreateFromMarkdownBlock(Paragraph);

                            }
                        }
                            
                    }

                  
                }

            }
        }



        public void GeneratePresentation()
        {

            // Add Slides
            foreach (var slide in _PresentationStructure.Slides)
            {
                SlideRange slideRange = _SlideManager
                    .CloneSlide(_Presentation, _Presentation.Slides[slide.TemplateSlide.Order], Locations.Location.Last);

                // Add Note content
                // slideRange.NotesPage



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
            foreach (SlideStructure slide in _TemplateStructure.Slides)
            {
                _SlideManager.DeleteSlide(_Presentation.Slides[1]);
            }
           



            this.SaveAs(Environment.CurrentDirectory + "/" + "output.pptx");
            this.Close();

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

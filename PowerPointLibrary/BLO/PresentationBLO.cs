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
        private readonly SlideZoneStructureBLO _SlideZoneStructureBLO;

        public PresentationBLO()
        {
            // Init Manager
            _ApplicationManager = new PowerPointApplicationManager();
            _PresentationManager = new PresentationManager();
            _SlideManager = new SlideManager();
            _TextRangeManager = new TextRangeManager();
            _ShapeManager = new ShapesManager();


            // Init BLO
            _TemplateStructureBLO = new PresentationStructureBLO();
            _TextStructureBLO = new TextStructureBLO();
            _CommentActionBLO = new CommentActionBLO();
            _SlideBLO = new SlideBLO(this);
            _SlideZoneStructureBLO = new SlideZoneStructureBLO();


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

                if (element is Microsoft.Toolkit.Parsers.Markdown.Blocks.CodeBlock code)
                {

                }

                if (element is Microsoft.Toolkit.Parsers.Markdown.Blocks.LinkReferenceBlock LinkReference)
                {

                }

                if (element is Microsoft.Toolkit.Parsers.Markdown.Blocks.ListBlock List)
                {
                    _SlideZoneStructureBLO.AddMarkdownBlockToSlideZone(this.CurrentSlide.CurrentZone, List);
                    this.CurrentSlide.CurrentZone.Text.Text += "\r";
                }

                if (element is Microsoft.Toolkit.Parsers.Markdown.Blocks.QuoteBlock Quote)
                {

                }

                if (element is Microsoft.Toolkit.Parsers.Markdown.Blocks.TableBlock Table)
                {

                }

                if (element is HeaderBlock header)
                {
                    string layout = "";
                    if (header.HeaderLevel == 1) layout = "Titre partie";
                    if (header.HeaderLevel >= 2) layout = "Titre et contenu";

                    _SlideBLO.AddSlide(layout);

                    SlideZoneStructure zoneTitle = this.CurrentSlide.CurrentZone;

                    if (zoneTitle != null)
                    {
                        if (zoneTitle.Text == null) zoneTitle.Text = new TextStructure();
                        _SlideZoneStructureBLO.AddMarkdownBlockToSlideZone(zoneTitle, header);
                    }
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
                            if (this.CurrentSlide.CurrentZone.Text == null)
                                this.CurrentSlide.CurrentZone.Text = new TextStructure();

                            // return à la ligne si une nouvelle paragraphe est ajouté
                            int count_befor = this.CurrentSlide.CurrentZone.Text.Text.Count();
                            _SlideZoneStructureBLO.AddMarkdownBlockToSlideZone(this.CurrentSlide.CurrentZone, Paragraph);
                            if (this.CurrentSlide.CurrentZone.Text.Text.Count() > count_befor)
                                this.CurrentSlide.CurrentZone.Text.Text += "\r";


                        }
                        else
                        {
                            if (this.CurrentSlide.AddToNotes)
                            {

                                //if (this.CurrentSlide.CurrentZone.Text == null) this.CurrentSlide.CurrentZone.Text = new TextStructure();
                                //_TextStructureBLO.CreateAndAddFromMarkdownBlock(this.CurrentSlide.CurrentZone.Text, Paragraph);

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

                Slide currentSlide = _Presentation.Slides[slideRange.SlideIndex];
                // Add Note content
                // slideRange.NotesPage



                foreach (var SlideZone in slide.SlideZones)
                {
                    // Add Text
                    if (SlideZone.Text != null)
                    {
                        Microsoft.Office.Interop.PowerPoint.Shape shape = slideRange.Shapes[SlideZone.Name];
                        //   shape.Fill.UserPicture( Environment.CurrentDirectory +  "/images/informatique.jpg");

                        if (SlideZone.Text != null)
                        {
                            _TextRangeManager.AddTextStructure(shape.TextFrame.TextRange, SlideZone.Text);
                        }

                        if (SlideZone.Image != null)
                        {

                            float imageHeight = 0;
                            float imageWidth = 0;
                            string file = Environment.CurrentDirectory + SlideZone.Image.Url;
                            using (var img = Image.FromFile(file))
                            {
                                imageHeight = img.Height;
                                imageWidth = img.Width;
                            }

                            float scale = Math.Min(shape.Width / imageWidth, shape.Height / imageHeight);

                            float scaledWidth = imageWidth * scale;
                            float scaledHeight = imageHeight * scale;


                            float left = (shape.Width - scaledWidth) / 2 + shape.Left;
                            float top = (shape.Height - scaledHeight) / 2 + shape.Top;

                            _ShapeManager.AddPicture(currentSlide, file, left, top, scaledWidth, scaledHeight);


                            // _ShapeManager.AddPicture(currentSlide, file, shape.Left, shape.Top, imageWidth, imageHeight); ;
                        }

                      
                       
                        

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

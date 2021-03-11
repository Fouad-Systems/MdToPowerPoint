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

        public PplArguments pplArguments;
        public PresentationBLO(PplArguments pplArguments)
        {
            this.pplArguments = pplArguments;

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


            this._TemplateStructure = _TemplateStructureBLO.LoadConfiguration(this.pplArguments.TemplateConfigurationPath);

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

                                if (this.CurrentSlide.Notes == null) this.CurrentSlide.Notes = new SlideZoneStructure();
                                if (this.CurrentSlide.Notes.Text == null) this.CurrentSlide.Notes.Text = new TextStructure();

                                _SlideZoneStructureBLO.AddMarkdownBlockToSlideZone(this.CurrentSlide.Notes, Paragraph);

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


                //currentSlide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.InsertAfter("This is a Test");
                //var ttt = currentSlide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange;

                if (slide.Notes != null && slide.Notes.Text != null)
                    _TextRangeManager.AddTextStructure(currentSlide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange, slide.Notes.Text);
                // slideRange.NotesPage



                foreach (var SlideZone in slide.SlideZones)
                {
                    // Add Text
                    if (SlideZone.Text != null)
                    {
                        Microsoft.Office.Interop.PowerPoint.Shape shape = slideRange.Shapes[SlideZone.Name];
                        //   shape.Fill.UserPicture( Environment.CurrentDirectory +  "/images/informatique.jpg");

                       
                        float shapeWidth = shape.Width;
                        float shapeHeight = shape.Height;
                        float shapeLeft = shape.Left;
                        float shapeTop = shape.Top;
                        if (SlideZone.Text != null)
                        {
                            _TextRangeManager.AddTextStructure(shape.TextFrame.TextRange, SlideZone.Text);
                        }

                        if (SlideZone.Images != null && SlideZone.Images.Count > 0)
                        {

                            foreach (var image in SlideZone.Images)
                            {

                                float imageHeight = 0;
                                float imageWidth = 0;
                                string file = Path.Combine(pplArguments.MdDocumentDirectoryPath, image.Url);
                                using (var img = Image.FromFile(file))
                                {
                                    imageHeight = img.Height;
                                    imageWidth = img.Width;
                                }

                                float scale = Math.Min(shapeWidth / imageWidth, shapeHeight / imageHeight);

                                float scaledWidth = imageWidth * scale;
                                float scaledHeight = imageHeight * scale;


                                float left = (shapeWidth - scaledWidth) / 2 + shapeLeft;
                                float top = (shapeHeight - scaledHeight) / 2 + shapeTop;

                                // _ShapeManager.AddPicture(currentSlide, file, left, top, scaledWidth, scaledHeight);
                                file = @"E:\formations\formation-git-github\src\images\10.jpg";
                                _ShapeManager.AddPicture(currentSlide, file, 10f,10f, 100f, 100f);
                            }




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




            this.SaveAs(this.pplArguments.OutPutFile);
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

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
            _SlideBLO = new SlideBLO(this);
            _SlideZoneStructureBLO = new SlideZoneStructureBLO();
            _TemplateStructureBLO = new TemplateStructureBLO(_PresentationStructure);
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


        public void CreatePresentationDataStructure(MarkdownDocument mdDocument)
        {

            // il faut d'abord, trouver le nombre des slides avec le nombre de type de contenue dans 
            // chaque slide

            // ensuite choisir la layout convenable pour chaque contenue 

            // ensuite read data frm mdDocument o PresentationDataStrucure



            foreach (var element in mdDocument.Blocks)
            {
                if (element is HeaderBlock header)
                {

                    if (header.HeaderLevel <= 2)
                    {
                        string layout = "";
                        if (header.HeaderLevel == 1) layout = "Titre session";
                        if (header.HeaderLevel >= 2) layout = "Titre contenu";

                        _SlideBLO.AddSlide(layout);
                        _SlideBLO.WriteToTitleZone();

                        SlideZoneStructure zoneTitle = this.CurrentSlide.CurrentZone;

                        if (zoneTitle != null)
                        {
                            if (zoneTitle.Text == null) zoneTitle.Text = new TextStructure();
                            _SlideZoneStructureBLO.AddMarkdownBlockToSlideZone(zoneTitle, header);
                        }
                    }
                    else
                    {
                        _SlideBLO.WriteToTextZone();
                        SlideZoneStructure zoneTitle = this.CurrentSlide.CurrentZone;
                        if (zoneTitle != null)
                            _SlideZoneStructureBLO.AddMarkdownBlockToSlideZone(zoneTitle, header);
                    }


                }


                if (this.CurrentSlide.UseSlideOrder != 0) continue;

                if (element is Microsoft.Toolkit.Parsers.Markdown.Blocks.CodeBlock code)
                {
                    if (this.CurrentSlide.AddToNotes)
                    {
                        _SlideBLO.AddNotes(code);

                    }
                    else
                    {

                        _SlideBLO.WriteToTextZone();


                        if (this.CurrentSlide.CurrentZone != null)
                        {
                            if (this.CurrentSlide.CurrentZone.Text == null)
                                this.CurrentSlide.CurrentZone.Text = new TextStructure();

                            // return à la ligne si une nouvelle paragraphe est ajouté
                            int count_befor = this.CurrentSlide.CurrentZone.Text.Text.Count();
                            _SlideZoneStructureBLO.AddMarkdownBlockToSlideZone(this.CurrentSlide.CurrentZone, code);
                            if (this.CurrentSlide.CurrentZone.Text.Text.Count() > count_befor)
                                this.CurrentSlide.CurrentZone.Text.Text += "\r";


                        }

                    }
                }

                if (element is Microsoft.Toolkit.Parsers.Markdown.Blocks.LinkReferenceBlock LinkReference)
                {
                }

                if (element is Microsoft.Toolkit.Parsers.Markdown.Blocks.ListBlock List)
                {
                    if (this.CurrentSlide.AddToNotes)
                    {
                        _SlideBLO.AddNotes(List);
                    }
                    else
                    {
                        _SlideZoneStructureBLO.AddMarkdownBlockToSlideZone(this.CurrentSlide.CurrentZone, List);
                        this.CurrentSlide.CurrentZone.Text.Text += "\r";
                    }

                }

                if (element is Microsoft.Toolkit.Parsers.Markdown.Blocks.QuoteBlock Quote)
                {
                }

                if (element is Microsoft.Toolkit.Parsers.Markdown.Blocks.TableBlock Table)
                {

                }


                if (element is ParagraphBlock Paragraph)
                {

                    // if paragraphe is action
                    if (Paragraph.Inlines[0].Type == MarkdownInlineType.Comment
                        && _CommentActionBLO.IsAction(Paragraph.Inlines[0].ToString()))
                    {

                        string comment = Paragraph.Inlines[0].ToString();

                        CommentAction commentAction = _CommentActionBLO.ParseComment(comment);

                        switch (commentAction.ActionType)
                        {
                            case CommentAction.ActionTypes.ChangeLayout:

                                this.CurrentSlide.IsLayoutChangedByAction = true;
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

                            case CommentAction.ActionTypes.Note:
                                _SlideBLO.StartWriteToNote();
                                break;
                            case CommentAction.ActionTypes.EndNote:
                                _SlideBLO.EndWriteToNote();
                                break;
                            case CommentAction.ActionTypes.Empty:
                                break;
                            case CommentAction.ActionTypes.UseSlide:
                                _SlideBLO.UseSlide(commentAction);
                                break;
                            case CommentAction.ActionTypes.GenerateLayout:
                                _GLayoutStructureBLO.GenerateSlideZone(this.CurrentSlide, commentAction.GLayoutStructure);
                                break;
                        }
                    }
                    else
                    {

                        if (this.CurrentSlide.AddToNotes)
                        {

                            _SlideBLO.AddNotes(Paragraph);



                        }
                        else
                        {

                            if (new ParagraphBlockBLO().IsImage(Paragraph))
                            {
                                _SlideBLO.WriteToImageZone();
                            }
                            else
                            {
                                _SlideBLO.WriteToTextZone();
                            }

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

                        }





                    }


                }

            }
        }

        public void GeneratePresentation()
        {
            this.PandocCommande("");

            //string output = process.StandardOutput.ReadToEnd();
           

            //string errors = process.StandardError.ReadToEnd();
            


            // Add Slides
            foreach (var slide in _PresentationStructure.Slides)
            {

                if (slide.UseSlideOrder != 0)
                {
                    // Use the slide from the file OutputfileName.slides.pptx

                    if (!File.Exists(pplArguments.UseSlideOutPutFile))
                    {
                        throw new PplException($"The file '{pplArguments.UseSlideOutPutFile}' doesn't exist");
                    }

                    Presentation PresentationSource = _PresentationManager
                        .OpenExistingPowerPointPresentation(_Application, pplArguments.UseSlideOutPutFile);

                    _SlideManager.CopySlideFromOtherPresentation(PresentationSource, slide.UseSlideOrder, _Presentation, _Presentation.Slides.Count);

                    continue;
                }


                // Clone Slide
                Slide currentSlide = null;
                SlideRange slideRange = null;
                if (slide.IsGenerated)
                {
                    int TitreContenuSldeOrder = _TemplateStructureBLO.GetSlide("Titre contenu").Order;
                    slideRange = _SlideManager
                   .CloneSlide(_Presentation, _Presentation.Slides[TitreContenuSldeOrder], Locations.Location.Last);
                    currentSlide = _Presentation.Slides[slideRange.SlideIndex];
                }
                else
                {
                    slideRange = _SlideManager
                   .CloneSlide(_Presentation, _Presentation.Slides[slide.TemplateSlide.Order], Locations.Location.Last);

                    currentSlide = _Presentation.Slides[slideRange.SlideIndex];
                }




                //currentSlide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.InsertAfter("This is a Test");
                //var ttt = currentSlide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange;

                if (slide.Notes != null && slide.Notes.Text != null)
                    _TextRangeManager.AddTextStructure(currentSlide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange, slide.Notes.Text);
                // slideRange.NotesPage

                if (slide.IsGenerated)
                {

           
                    float resolution_with = currentSlide.Master.Width;
                    float resolution_Height = currentSlide.Master.Height;
                    float ratio = resolution_with / 1920;
                    ratio = resolution_Height / 1080;

                    Microsoft.Office.Interop.PowerPoint.Shape contenu_shape = currentSlide.Shapes["contenu"];

                    foreach (var SlideZone in slide.GeneratedSlideZones)
                    {

                      
                        if(SlideZone.Name == "Title" || SlideZone.Name =="Titre")
                        {
                            Microsoft.Office.Interop.PowerPoint.Shape titleShape = slideRange.Shapes[SlideZone.Name];
                         

                            if (SlideZone.Text != null && !string.IsNullOrEmpty(SlideZone.Text.Text))
                            {

                                _TextRangeManager.AddTextStructure(titleShape.TextFrame.TextRange, SlideZone.Text);

                            }

                            continue;
                        }
                       

                        

                        if (SlideZone.Text != null && !string.IsNullOrEmpty(SlideZone.Text.Text))
                        {

                            //var range = contenu_shape.Duplicate();
                            //Microsoft.Office.Interop.PowerPoint.Shape shape1 = range[1];
                            //shape1.Left = 100;
                            //shape1.Top = 200;
                            //shape1.Width = 400;
                            //shape1.Height = 300;



                            var shape = contenu_shape.Duplicate()[1];

                            shape.Width = SlideZone.Width * ratio;
                            shape.Height = SlideZone.Height * ratio;
                            shape.Top = SlideZone.Top * ratio;
                            shape.Left = SlideZone.Left * ratio;
                            // shape.Fill.BackColor.RGB = Color.Pink.ToArgb();

                           // shape.Visible = MsoTriState.msoCTrue;

                            _TextRangeManager.AddTextStructure(shape.TextFrame.TextRange, SlideZone.Text);

                            if (!SlideZone.ContentTypes.Contains(Entities.Enums.ContentTypes.Title))
                            {
                                shape.AnimationSettings.EntryEffect = PpEntryEffect.ppEffectAppear;
                            }

                        }

                        if (SlideZone.Images != null && SlideZone.Images.Count > 0)
                        {


                            

                            //var shape = contenu_shape.Duplicate()[1];

                            //shape.Width = SlideZone.Width * ratio;
                            //shape.Height = SlideZone.Height * ratio;
                            //shape.Top = SlideZone.Top * ratio;
                            //shape.Left = SlideZone.Left * ratio;
                            //shape.Fill.BackColor.RGB = Color.Green.ToArgb();


                            // shape.Visible = MsoTriState.msoTrue;


                            float shapeWidth = SlideZone.Width * ratio;
                            float shapeHeight = SlideZone.Height * ratio;
                            float shapeLeft = SlideZone.Left * ratio;
                            float shapeTop = SlideZone.Top * ratio;

                            // il faut supprimer "shape" si non AddPicture va remplater le shpae par image 
                            // sans prendre en considération with et height

                            // shape.Delete(); if yout delet a shape PowerPoit will rename the other shape
                            // add empty string for AddPicture to Word Correctly with with et and height
                            // shape.TextFrame.TextRange.Text = "aaaaaaaa a ";

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

                                Microsoft.Office.Interop.PowerPoint.Shape image_shape = _ShapeManager
                                    .AddPicture(currentSlide, file, left, top, scaledWidth, scaledHeight);

                                // contenu_shape.Visible = MsoTriState.msoFalse;

                                //Microsoft.Office.Interop.PowerPoint.Shape image_shape = _ShapeManager
                                //    .AddPicture(currentSlide, file, shape.Left, shape.Top, shape.Width, shape.Height);

                                //  image_shape.AnimationSettings.AnimationOrder = 1;
                                image_shape.AnimationSettings.EntryEffect = PpEntryEffect.ppEffectAppear;
                                //file = @"E:\formations\formation-git-github\src\images\10.jpg";
                                // _ShapeManager.AddPicture(currentSlide, file, 10f,10f, 100f, 100f);
                            }


                            // _ShapeManager.AddPicture(currentSlide, file, shape.Left, shape.Top, imageWidth, imageHeight); ;
                        }

                      
                    }

                    contenu_shape.TextFrame.TextRange.Text = " ";
                }
                else
                {
                    foreach (var SlideZone in slide.SlideZones)
                    {


                        //List<string> zones = new List<string>();
                        //foreach (Microsoft.Office.Interop.PowerPoint.Shape item in slideRange.Shapes)
                        //{
                        //    zones.Add(item.Name);
                        //}

                        Microsoft.Office.Interop.PowerPoint.Shape shape = slideRange.Shapes[SlideZone.Name];
                        //   shape.Fill.UserPicture( Environment.CurrentDirectory +  "/images/informatique.jpg");

                        if (SlideZone.Text != null && !string.IsNullOrEmpty(SlideZone.Text.Text))
                        {

                            Thread thread = new Thread(() => Clipboard.SetText(SlideZone.Text.Text, TextDataFormat.UnicodeText));
                            thread.SetApartmentState(ApartmentState.STA); //Set the thread to STA
                            thread.Start();
                            thread.Join();

                         

                            //  slideRange.Shapes.AddShape(MsoAutoShapeType.msoShapeLineCallout1, 1, 1, 1, 1);
                            // slideRange.Shapes.PasteSpecial(PpPasteDataType.ppPasteHTML);

                            //shape.TextFrame
                            //// Clipboard.SetText();
                            //// shape.TextFrame.TextRange.Paste();
                            shape.TextFrame.TextRange.PasteSpecial(PpPasteDataType.ppPasteText) ;

                         

                            //  _TextRangeManager.AddTextStructure(shape.TextFrame.TextRange, SlideZone.Text);

                            if (!SlideZone.ContentTypes.Contains(Entities.Enums.ContentTypes.Title))
                            {
                                shape.AnimationSettings.EntryEffect = PpEntryEffect.ppEffectAppear;
                            }

                        }

                        if (SlideZone.Images != null && SlideZone.Images.Count > 0)
                        {
                            float shapeWidth = shape.Width;
                            float shapeHeight = shape.Height;
                            float shapeLeft = shape.Left;
                            float shapeTop = shape.Top;

                            // il faut supprimer "shape" si non AddPicture va remplater le shpae par image 
                            // sans prendre en considération with et height

                            // shape.Delete(); if yout delet a shape PowerPoit will rename the other shape
                            // add empty string for AddPicture to Word Correctly with with et and height
                            // shape.TextFrame.TextRange.Text = "aaaaaaaa a ";

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

                                Microsoft.Office.Interop.PowerPoint.Shape image_shape = _ShapeManager
                                    .AddPicture(currentSlide, file, left, top, scaledWidth, scaledHeight);

                                //  image_shape.AnimationSettings.AnimationOrder = 1;
                                image_shape.AnimationSettings.EntryEffect = PpEntryEffect.ppEffectAppear;
                                //file = @"E:\formations\formation-git-github\src\images\10.jpg";
                                // _ShapeManager.AddPicture(currentSlide, file, 10f,10f, 100f, 100f);
                            }


                            // _ShapeManager.AddPicture(currentSlide, file, shape.Left, shape.Top, imageWidth, imageHeight); ;
                        }
                    }

                }



            }

            // Delete Template Slide
            foreach (SlideStructure slide in this._PresentationStructure._TemplateStructure.Slides)
            {
                _SlideManager.DeleteSlide(_Presentation.Slides[1]);
            }




            this.SaveAs(this.pplArguments.OutPutFile);
            this.Close();

        }



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

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
    public class GeneratePresentationBLO
    {
        public PresentationStructure _PresentationStructure;
        public Microsoft.Office.Interop.PowerPoint.Application _Application;
        public Presentation _Presentation;
        public PplArguments pplArguments;

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


        public GeneratePresentationBLO(
            Microsoft.Office.Interop.PowerPoint.Application _Application,
            Presentation _Presentation,
            PresentationStructure _PresentationStructure, 
            PplArguments pplArguments)
        {
            this._Application = _Application;
            this._Presentation = _Presentation;
            this.pplArguments = pplArguments;
            this._PresentationStructure = _PresentationStructure;


            // Init Manager
            _ApplicationManager = new PowerPointApplicationManager();
            _PresentationManager = new PresentationManager();
            _SlideManager = new SlideManager();
            _TextRangeManager = new TextRangeManager();
            _ShapeManager = new ShapesManager();


            // Init BLO
            _PresentationStructureBLO = new PresentationStructureBLO(this._PresentationStructure);
            _TextStructureBLO = new TextStructureBLO();
            _CommentActionBLO = new CommentActionBLO();
            _SlideBLO = new SlideBLO(_PresentationStructure);
            _SlideZoneStructureBLO = new SlideZoneStructureBLO();
            _TemplateStructureBLO = new TemplateStructureBLO(_PresentationStructure);
        }

        public void GeneratePresentation()
        {
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


                        if (SlideZone.Name == "Title" || SlideZone.Name == "Titre")
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

                            _TextRangeManager.AddTextStructure(shape.TextFrame.TextRange, SlideZone.Text);

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
 
        }

    }
}

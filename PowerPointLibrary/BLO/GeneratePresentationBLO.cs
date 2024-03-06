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
    /// Generare a presentation from PresentationStructure
    /// </summary>
    public class GeneratePresentationBLO
    {
        #region attributes 

        public Microsoft.Office.Interop.PowerPoint.Application _Application;
        public  PresentationStructure _PresentationStructure;
        public  Presentation _Presentation;
        public  TopptArguments pplArguments;

        private PowerPointApplicationManager _ApplicationManager;
        private PresentationManager _PresentationManager;
        private SlideManager _SlideManager;
        private ShapesManager _ShapeManager;
        private TextRangeManager _TextRangeManager;

        private PresentationStructureBLO _PresentationStructureBLO;
        private TemplateBLO _TemplateStructureBLO;
        private CommentActionBLO _CommentActionBLO;
        private SlideBLO _SlideBLO;
        private SlideZoneBLO _SlideZoneStructureBLO;
        private LayoutGeneratorBLO _GLayoutStructureBLO;

        #endregion

        public string ImagePatheFile(string ImageURL)
        {
  
            if (ImageURL.StartsWith("../"))
            {
                // TODO : Il faut ajouter ça dans le fichier de configuration
                // si le chemin de l'image est relative, 
                // dans mon chaine de production dans Jekyll, je dépose les articles à l'intérieur d'un document
                // dans la collection _Chapitres, pour corriger les chemins des images il faut ajoutert : "../"
                ImageURL = "../" + ImageURL;
            }
                
            else
            {
                // Pourquoi ajouter "../../" ?
                // l'image est dans _Chapitre/Comprendre-ordinateur/informatique.md  : ./image/photo.png
                // l'image est dans ../.././image/photo.png
                ImageURL = "../../" + ImageURL;
            }




            string imageFilePath = Path.Combine(pplArguments.MdDocumentDirectoryPath, ImageURL);
            return imageFilePath;
        }

        public GeneratePresentationBLO(
            Microsoft.Office.Interop.PowerPoint.Application _Application,
            Presentation _Presentation,
            PresentationStructure _PresentationStructure, 
            TopptArguments pplArguments)
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
            _CommentActionBLO = new CommentActionBLO();
            _SlideBLO = new SlideBLO(_PresentationStructure);
            _SlideZoneStructureBLO = new SlideZoneBLO();
            _TemplateStructureBLO = new TemplateBLO(_PresentationStructure);
        }


        /// <summary>
        /// Create the powserpoint presentation from the PresentationStructure instance
        /// </summary>
        public void GeneratePresentation()
        {

            foreach (var slide in _PresentationStructure.Slides)
            {
                // Use the slide from the presentation : OutputfileName.slide.pptx
                if (slide.UseSlideOrder != 0)
                {
                    this.UseSlide(slide);
                    continue;
                }

                Slide currentSlide = this.CreateSlide(slide);
                this.AddNoteToSlide(slide, currentSlide);

                // Calculate ration beetween VBA Resulution and powerPoint Resoltion
                float resolution_with = currentSlide.Master.Width;
                float resolution_Height = currentSlide.Master.Height;
                float ratio = resolution_with / 1920;
                ratio = resolution_Height / 1080;

                if (slide.IsGenerated)
                {
                    this.CreateGeneratedSlide(slide, currentSlide, ratio);
                }
                else
                {
                    this.CreateSlide(slide, currentSlide, ratio);
                 
                }

            }

            // Delete Template Slide
            foreach (SlideStructure slide in this._PresentationStructure._TemplateStructure.Slides)
            {
                _SlideManager.DeleteSlide(_Presentation.Slides[1]);
            }
 
        }

        private void CreateSlide(SlideStructure slide, Slide currentSlide, float ratio)
        {
            foreach (var SlideZone in slide.SlideZones)
            {
                // Find the shape 
                Microsoft.Office.Interop.PowerPoint.Shape shape = currentSlide.Shapes[SlideZone.Name];

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
                    SlideZoneStructure shapeZone = new SlideZoneStructure();

                    shapeZone.Width = shape.Width;
                    shapeZone.Height = shape.Height;
                    shapeZone.Left = shape.Left;
                    shapeZone.Top = shape.Top;

                    foreach (var image in SlideZone.Images)
                    {
                        string imageFilePath = this.ImagePatheFile(image.Url);

                        // Center the image in the shape : calculate the new dimension
                        SlideZoneStructure ImageDimension = this.CenterImageInShape(shapeZone, ratio, imageFilePath);


                        Microsoft.Office.Interop.PowerPoint.Shape image_shape = _ShapeManager
                            .AddPicture(currentSlide,
                            imageFilePath,
                            ImageDimension.Left,
                            ImageDimension.Top,
                            ImageDimension.Width,
                            ImageDimension.Height);

                        // Create animation
                        image_shape.AnimationSettings.EntryEffect = PpEntryEffect.ppEffectAppear;
                    }

                }
            }

        }

        private void CreateGeneratedSlide(SlideStructure slide, Slide currentSlide, float ratio)
        {
            // Find the shape 
            Microsoft.Office.Interop.PowerPoint.Shape contenu_shape = currentSlide.Shapes["contenu"];

          

            // Create a shape for each SlideZone
            foreach (var SlideZone in slide.GeneratedSlideZones)
            {
                // Add content to Title zone
                if (SlideZone.Name == "Title" || SlideZone.Name == "Titre")
                {
                    Microsoft.Office.Interop.PowerPoint.Shape titleShape = currentSlide.Shapes[SlideZone.Name];
                    if (SlideZone.Text != null && !string.IsNullOrEmpty(SlideZone.Text.Text))
                    {
                        _TextRangeManager.AddTextStructure(titleShape.TextFrame.TextRange, SlideZone.Text);
                    }
                    continue;
                }

                // Insert image in current zone
                if (SlideZone.Images != null && SlideZone.Images.Count > 0)
                {

                    foreach (var image in SlideZone.Images)
                    {
                       

                        string imageFilePath = this.ImagePatheFile( image.Url);

                       
                        // Center the image in the shape : calculate the new dimension
                        SlideZoneStructure ImageDimension = this.CenterImageInShape(SlideZone, ratio, imageFilePath);

                        // Insert image 
                        Microsoft.Office.Interop.PowerPoint.Shape image_shape = _ShapeManager
                            .AddPicture(currentSlide,
                            imageFilePath,
                            ImageDimension.Left,
                            ImageDimension.Top,
                            ImageDimension.Width,
                            ImageDimension.Height);

                       


                        // Create annimation
                        image_shape.AnimationSettings.EntryEffect = PpEntryEffect.ppEffectAppear;
   
                    }

                    continue;
                }


                // Add text to current zone
                if (SlideZone.Text != null && !string.IsNullOrEmpty(SlideZone.Text.Text))
                {
                    // Creae a shape
                    var shape = contenu_shape.Duplicate()[1];
                    shape.Width = SlideZone.Width * ratio;
                    shape.Height = SlideZone.Height * ratio;
                    shape.Top = SlideZone.Top * ratio;
                    shape.Left = SlideZone.Left * ratio;

                    // Add Data to shape
                    _TextRangeManager.AddTextStructure(shape.TextFrame.TextRange, SlideZone.Text);

                    // Create annimation  
                    if (!SlideZone.ContentTypes.Contains(Entities.Enums.ContentTypes.Title))
                    {
                        shape.AnimationSettings.Animate = MsoTriState.msoCTrue;
                        shape.AnimationSettings.TextLevelEffect = PpTextLevelEffect.ppAnimateByFifthLevel;

                        shape.AnimationSettings.TextUnitEffect = PpTextUnitEffect.ppAnimateByParagraph;
                      //  shape.AnimationSettings.EntryEffect = PpEntryEffect.ppEffectAppear;
                  

                    }

                }

               

            }

            // hide the shape "contnue", you mustn't delete it.
            contenu_shape.TextFrame.TextRange.Text = " ";


            // Set annimation time = 0 

            for (int i = 1; i <= currentSlide.TimeLine.MainSequence.Count; i++)
            {
                currentSlide.TimeLine.MainSequence[i].Timing.Duration = 0;
            }
        }

        private SlideZoneStructure CenterImageInShape(SlideZoneStructure slideZone, float ratio, string imageFilePath)
        {
            SlideZoneStructure ImageDimension = new SlideZoneStructure();
            
            float imageHeight = 0;
            float imageWidth = 0;

        

            float shapeWidth = slideZone.Width * ratio;
            float shapeHeight = slideZone.Height * ratio;
            float shapeLeft = slideZone.Left * ratio;
            float shapeTop = slideZone.Top * ratio;

            using (var img = Image.FromFile(imageFilePath))
            {
                imageHeight = img.Height;
                imageWidth = img.Width;
            }
            float scale = Math.Min(shapeWidth / imageWidth, shapeHeight / imageHeight);

            ImageDimension.Width = imageWidth * scale;
            ImageDimension.Height = imageHeight * scale;
            ImageDimension.Left = (shapeWidth - ImageDimension.Width) / 2 + shapeLeft;
            ImageDimension.Top = (shapeHeight - ImageDimension.Height) / 2 + shapeTop;

            return ImageDimension;
        }

        private void CreateSlideZoneInCurrentSlide(SlideZoneStructure SlideZone, Slide currentSlide)
        {
        }

        private void AddNoteToSlide(SlideStructure slide, Slide currentSlide )
        {
            // Add Notes to Slide
            if (slide.Notes != null && slide.Notes.Text != null)
                _TextRangeManager.AddTextStructure(currentSlide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange, slide.Notes.Text);

        }

        /// <summary>
        /// 
        /// </summary>
        private Slide CreateSlide(SlideStructure slide)
        {
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
            return currentSlide;
        }

        /// <summary>
        ///  // Use the slide from the presentation : OutputfileName.slide.pptx
        /// </summary>
        /// <param name="useSlideOrder"></param>
        private void UseSlide(SlideStructure slide)
        {
            // Use the slide from the file OutputfileName.slides.pptx
            if (!File.Exists(pplArguments.UseSlideOutPutFile))
            {
                throw new PplException($"The file '{pplArguments.UseSlideOutPutFile}' doesn't exist");
            }

            Presentation PresentationSource = _PresentationManager
                .OpenExistingPowerPointPresentation(_Application, pplArguments.UseSlideOutPutFile);

            _SlideManager.CopySlideFromOtherPresentation(PresentationSource, slide.UseSlideOrder, _Presentation, _Presentation.Slides.Count);

        }
    }
}

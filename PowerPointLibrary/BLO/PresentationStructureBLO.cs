using Microsoft.Toolkit.Parsers.Markdown;
using Microsoft.Toolkit.Parsers.Markdown.Blocks;
using Newtonsoft.Json;
using PowerPointLibrary.Entities;
using PowerPointLibrary.Exceptions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.BLO
{
    /// <summary>
    /// Create presentationStructure from Markdown structure
    /// </summary>
    public class PresentationStructureBLO
    {

        private PresentationStructure _PresentationStructure;

        //private PresentationStructureBLO _PresentationStructureBLO;
        private TemplateStructureBLO _TemplateStructureBLO;
        private TextStructureBLO _TextStructureBLO;
        private CommentActionBLO _CommentActionBLO;
        private SlideBLO _SlideBLO;
        private SlideZoneStructureBLO _SlideZoneStructureBLO;
        private GLayoutStructureBLO _GLayoutStructureBLO;


        public PresentationStructureBLO(PresentationStructure presentationStructure)
        {
            _PresentationStructure = presentationStructure;

            // Init BLO
            _TextStructureBLO = new TextStructureBLO();
            _CommentActionBLO = new CommentActionBLO();
            _SlideBLO = new SlideBLO(_PresentationStructure);
            _SlideZoneStructureBLO = new SlideZoneStructureBLO();
            _TemplateStructureBLO = new TemplateStructureBLO(_PresentationStructure);
            _GLayoutStructureBLO = new GLayoutStructureBLO();
        }


        public void CreatePresentationDataStructure(MarkdownDocument mdDocument)
        {
            // Amélioration
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
                        _SlideBLO.WriteToTextZone();
                        if (this.CurrentSlide.CurrentZone != null) {
                            _SlideZoneStructureBLO.AddMarkdownBlockToSlideZone(this.CurrentSlide.CurrentZone, List);
                            this.CurrentSlide.CurrentZone.Text.Text += "\r";
                        }
                           
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
                                _SlideBLO.NewSlide(commentAction.Layout);
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
                            case CommentAction.ActionTypes.NewZone:
                                _SlideBLO
                                   .NewZone(this.CurrentSlide);
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


        public static void CreateTemplateStructureExemple()
        {

            PresentationStructure templateStructure = new PresentationStructure();

            SlideStructure slide1 = new SlideStructure();
            slide1.Name = "Slide 1";
            slide1.Order = 1;
            slide1.SlideZones.Add(new SlideZoneStructure() { Name = "zone1" });
            slide1.SlideZones.Add(new SlideZoneStructure() { Name = "zone2" });
            //slide1.ContentTypes.Add(Entities.Enums.ContentTypes.Title);
            //slide1.ContentTypes.Add(Entities.Enums.ContentTypes.Text);


            SlideStructure slide2 = new SlideStructure();
            slide2.Name = "Slide 2";
            slide2.Order = 1;
            slide2.SlideZones.Add(new SlideZoneStructure() { Name = "zone1" });
            slide2.SlideZones.Add(new SlideZoneStructure() { Name = "zone2" });
            slide2.SlideZones.Add(new SlideZoneStructure() { Name = "zone3" });

            templateStructure.Slides.Add(slide1);
            templateStructure.Slides.Add(slide2);

            File.WriteAllText(@"exemple-template.json", JsonConvert.SerializeObject(templateStructure));


        }

        public PresentationStructure LoadConfiguration(string templateName)
        {


            string code = File.ReadAllText(templateName);

            var obj = JsonConvert.DeserializeObject(code, typeof(PresentationStructure));
            PresentationStructure templateStructure = obj as PresentationStructure;

            // Calculate Zone Order 

            foreach (var slide in templateStructure.Slides)
            {
                int order = 1;
                foreach (var slideZone in slide.SlideZones)
                {
                    slideZone.Order = order++;
                }

            }

            return templateStructure;
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

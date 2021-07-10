using Microsoft.Toolkit.Parsers.Markdown.Blocks;
using PowerPointLibrary.Entities;
using PowerPointLibrary.Entities.Enums;
using PowerPointLibrary.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.BLO
{
    public class SlideBLO
    {
        public static string ZoneTitleName = "Titre";
        private SlideZoneBLO _SlideZoneStructureBLO;
        private TemplateBLO _TemplateStructureBLO;
        private PresentationStructure _PresentationStructure;
        private LayoutGeneratorBLO _LayoutGeneratorBLO;
        public SlideBLO(PresentationStructure _PresentationStructure)
        {
            this._PresentationStructure = _PresentationStructure;
            _SlideZoneStructureBLO = new SlideZoneBLO();
            _TemplateStructureBLO = new TemplateBLO(_PresentationStructure);
            _LayoutGeneratorBLO = new LayoutGeneratorBLO();
        }

        public SlideStructure CurrentSlide
        {
            get
            {
                return _PresentationStructure.CurrentSlide;
            }
        }

        #region FindZone

        public static SlideZoneStructure GetTitleZone(SlideStructure slideStructure)
        {
            if (slideStructure.IsGenerated)
            {
                return slideStructure.GeneratedSlideZones.Where(z => z.ContentTypes.Contains(Entities.Enums.ContentTypes.Title))
                     .FirstOrDefault();
            }

            return slideStructure.SlideZones.Where(z => z.ContentTypes.Contains(Entities.Enums.ContentTypes.Title))
                   .FirstOrDefault();

        }

        public static SlideZoneStructure GetTitleZoneFromSlideZones(SlideStructure slideStructure)
        {

            return slideStructure.SlideZones.Where(z => z.ContentTypes.Contains(Entities.Enums.ContentTypes.Title))
                   .FirstOrDefault();

        }

        #endregion

        #region Apply CommentAction

        public CommentAction ApplyAction(string comment)
        {
            CommentActionBLO commentActionBLO = new CommentActionBLO();
            CommentAction commentAction = commentActionBLO.ParseComment(comment);
            this.ApplyAction(commentAction);
            return commentAction;
        }

        public void ApplyAction(CommentAction commentAction)
        {
            switch (commentAction.ActionType)
            {
                case CommentAction.ActionTypes.ChangeLayout:

                    this.CurrentSlide.IsLayoutChangedByAction = true;
                    this.ChangeLayout(this.CurrentSlide, commentAction.Layout);
                    break;
                case CommentAction.ActionTypes.ChangeZone:
                    this.ChangeCurrentZone(this.CurrentSlide, commentAction.ZoneName);
                    break;
                case CommentAction.ActionTypes.NewSlide:
                    this.NewSlide(commentAction.Layout);
                    break;
                case CommentAction.ActionTypes.Note:
                    this.StartWriteToNote();
                    break;
                case CommentAction.ActionTypes.EndNote:
                    this.EndWriteToNote();
                    break;
                case CommentAction.ActionTypes.Empty:
                    break;
                case CommentAction.ActionTypes.UseSlide:
                    this.UseSlide(commentAction);
                    break;
                case CommentAction.ActionTypes.GenerateLayout:
                    _LayoutGeneratorBLO.GenerateSlideZone(this.CurrentSlide, commentAction.GLayoutStructure);
                    break;
                case CommentAction.ActionTypes.NewZone:
                    this
                       .NewZone(this.CurrentSlide);
                    break;
            }
        }

        public void ChangeLayout(SlideStructure slideStructure, string layout)
        {
            slideStructure.Layout = layout;

            var oldSlideZones = slideStructure
                .SlideZones.Select(z => z.Clone() as SlideZoneStructure).ToList();

            var TemplateSlide = _TemplateStructureBLO.GetSlide(layout);

            slideStructure.SlideZones = TemplateSlide.SlideZones.Select(s => s.Clone() as SlideZoneStructure).ToList();
            slideStructure.TemplateSlide = TemplateSlide;

            // Copy old zone to the new layout, if old zone exist in the new layout
            foreach (var oldSlideZone in oldSlideZones)
            {

                for (int i = 0; i < slideStructure.SlideZones.Count; i++)
                {
                    if (slideStructure.SlideZones[i].Name == oldSlideZone.Name)
                    {
                        slideStructure.SlideZones[i] = oldSlideZone;
                        slideStructure.Order = i + 1;
                    }
                }
            }

            if (slideStructure.SlideZones.Count > 0)
                slideStructure.CurrentZone = slideStructure.SlideZones.First();
            else
                slideStructure.CurrentZone = null;
        }

        public void ChangeCurrentZone(SlideStructure currentSlide, string zoneName)
        {
            SlideZoneStructure CurrentZone = currentSlide.SlideZones.Where(z => z.Name == zoneName).FirstOrDefault();
            if (CurrentZone == null)
            {
                string msg = $"The zone name {zoneName} doesn't exist";
                throw new PowerPointLibrary.Exceptions.PplException(msg);
            }
            currentSlide.CurrentZone = CurrentZone;

        }

        public void NewZone(SlideStructure currentSlide)
        {

            if (this.CurrentSlide.IsGenerated)
            {
                int CurrentZoneOrder = this.CurrentSlide.GeneratedSlideZones.Where(z => !z.IsEmpty()).Count();

                if (this.CurrentSlide.CurrentZone != null) CurrentZoneOrder = this.CurrentSlide.CurrentZone.Order;

                this.CurrentSlide.CurrentZone = this.CurrentSlide
              .GeneratedSlideZones.Where(s => s.Order > CurrentZoneOrder).Where(z => z.ContentTypes.Contains(Entities.Enums.ContentTypes.Text))
              .FirstOrDefault();
            }
            else
            {
                int CurrentZoneOrder = this.CurrentSlide.SlideZones.Where(z => !z.IsEmpty()).Count();

                if (this.CurrentSlide.CurrentZone != null) CurrentZoneOrder = this.CurrentSlide.CurrentZone.Order;

                this.CurrentSlide.CurrentZone = this.CurrentSlide
              .SlideZones.Where(s => s.Order > CurrentZoneOrder).Where(z => z.ContentTypes.Contains(Entities.Enums.ContentTypes.Text))
              .FirstOrDefault();
            }

           


          

        }

        public void AddSlide(string Layout)
        {
            // Add Template Zone to Slide
            var TemplateSlide = _TemplateStructureBLO.GetSlide(Layout);

            SlideStructure slideStructure = new SlideStructure();
            _PresentationStructure.Slides.Add(slideStructure);

            slideStructure.Name = "Slide" + _PresentationStructure.Slides.Count;
            slideStructure.Layout = Layout;
            slideStructure.TemplateSlide = TemplateSlide;
            slideStructure.SlideZones = TemplateSlide.SlideZones.Select(s => s.Clone() as SlideZoneStructure).ToList();


        }

        public void NewSlide(string Layout)
        {
            var TitleZone = GetTitleZone(this.CurrentSlide);

        
            this.AddSlide(Layout);

            var CurrentTitleZone = GetTitleZone(this.CurrentSlide);
            TitleZone.Clone(CurrentTitleZone);

          

           


        }

        public void UseSlide(CommentAction commentAction)
        {
            this.CurrentSlide.UseSlideOrder = commentAction.UseSlideOrder;
        }

        #endregion

        #region FindZone
        public void FindZoneForTitle()
        {
            int CurrentZoneOrder = 0;
            if (this.CurrentSlide.CurrentZone != null) CurrentZoneOrder = this.CurrentSlide.CurrentZone.Order;


            this.CurrentSlide.CurrentZone = this.CurrentSlide
                .SlideZones.Where(s => s.Order > CurrentZoneOrder).Where(z => z.ContentTypes.Contains(Entities.Enums.ContentTypes.Title))
                .FirstOrDefault();
        }

        public void FindZoneForTexte()
        {
            // Change if we are in zone title or image zone

            if (this.CurrentSlide.IsGenerated == false)
            {
                if (
                    this.CurrentSlide.CurrentZone == null 
                    || this.CurrentSlide.CurrentZone.IsImage()
                    || this.CurrentSlide.CurrentZone.IsTitle()
                   
                    )
                {

                    int CurrentZoneOrder = this.CurrentSlide.SlideZones.Where(z => !z.IsEmpty()).Count();

                    if (this.CurrentSlide.CurrentZone != null) CurrentZoneOrder = this.CurrentSlide.CurrentZone.Order;

                    this.CurrentSlide.CurrentZone = this.CurrentSlide
                  .SlideZones.Where(s => s.Order > CurrentZoneOrder).Where(z => z.ContentTypes.Contains(Entities.Enums.ContentTypes.Text))
                  .FirstOrDefault();
                }
            }
            else
            {
                // Change if we are in zone title or image zone
                if (
                    this.CurrentSlide.CurrentZone == null
                    || this.CurrentSlide.CurrentZone.IsImage()
                    || this.CurrentSlide.CurrentZone.IsTitle()

                    )
                {
                    int CurrentZoneOrder = this.CurrentSlide.GeneratedSlideZones.Where(z => !z.IsEmpty()).Count();
                    if (this.CurrentSlide.CurrentZone != null) CurrentZoneOrder = this.CurrentSlide.CurrentZone.Order;

                    this.CurrentSlide.CurrentZone = this.CurrentSlide
                  .GeneratedSlideZones

                  .Where(s => s.Order > CurrentZoneOrder)
                  .Where(z => z.ContentTypes.Contains(Entities.Enums.ContentTypes.Text))
                  .FirstOrDefault();
                }
            }
              



        }

        public void FindZoneForImage()
        {
          


            if (this.CurrentSlide.IsGenerated == false)
            {
                int CurrentZoneOrder = this.CurrentSlide.SlideZones.Where(z => !z.IsEmpty()).Count();
                if (this.CurrentSlide.CurrentZone != null) CurrentZoneOrder = this.CurrentSlide.CurrentZone.Order;


                this.CurrentSlide.CurrentZone = this.CurrentSlide
             .SlideZones.Where(s => s.Order > CurrentZoneOrder)
             .Where(z => z.ContentTypes.Contains(Entities.Enums.ContentTypes.Image))
             .FirstOrDefault();
            }
            else
            {
                int CurrentZoneOrder = this.CurrentSlide.GeneratedSlideZones.Where(z => !z.IsEmpty()).Count();
                if (this.CurrentSlide.CurrentZone != null) CurrentZoneOrder = this.CurrentSlide.CurrentZone.Order;


                this.CurrentSlide.CurrentZone = this.CurrentSlide
            .GeneratedSlideZones.Where(s => s.Order > CurrentZoneOrder)
            .Where(z => z.ContentTypes.Contains(Entities.Enums.ContentTypes.Image))
            .FirstOrDefault();

            }



            //// Change layout to another layout that have a free space for the new image
            //if (this.CurrentSlide.CurrentZone == null && !this.CurrentSlide.IsLayoutChangedByAction)
            //{
            //    var ImageLayout = _PresentationBLO
            //        ._TemplateStructure
            //        .Slides
            //        .Where(s => s.SlideZones
            //           .Where(z => 

            //              z.ContentTypes.Contains(Entities.Enums.ContentTypes.Image) 
            //             ||

            //             z.ContentTypes.Contains(Entities.Enums.ContentTypes.Text)

            //           )
            //           .Count() >= 2

            //        )
            //       .Where(ss => ss.SlideZones.Where(zz => zz.ContentTypes.Contains(Entities.Enums.ContentTypes.Image)).FirstOrDefault() != null )

            //       .FirstOrDefault();

            //    if (ImageLayout != null)
            //    {
            //        this.ChangeLayout(this.CurrentSlide, ImageLayout.Name);

            //        this.CurrentSlide.CurrentZone = this.CurrentSlide
            // .SlideZones.Where(s => s.Order > CurrentZoneOrder).Where(z => z.ContentTypes.Contains(Entities.Enums.ContentTypes.Image))
            // .FirstOrDefault();
            //    }

            //}

        }

       

        #endregion

        #region Add Note
        public void StartWriteToNote()
        {
            CurrentSlide.AddToNotes = true;
        }

        public void EndWriteToNote()
        {
            CurrentSlide.AddToNotes = false;
        }

        /// <summary>
        /// Add paragraphe to Note zone
        /// </summary>
        /// <param name="paragraph"></param>
        public void AddNotes(MarkdownBlock paragraph, int contentNumber)
        {
            if (this.CurrentSlide.Notes == null) this.CurrentSlide.Notes = new SlideZoneStructure();
            if (this.CurrentSlide.Notes.Text == null) this.CurrentSlide.Notes.Text = new TextStructure();

            this.CurrentSlide.Notes.Text.Text += $"{contentNumber} |";

            _SlideZoneStructureBLO.AddMarkdownBlockToSlideZone(this.CurrentSlide.Notes, paragraph);
            this.CurrentSlide.Notes.Text.Text += "\r";
        }

        public void AddExplicationNotes(MarkdownBlock markdownBlock, int contentNumber)
        {
            MarkdownBlockBLO markdownBlockBLO = new MarkdownBlockBLO();

            if (this.CurrentSlide.Notes == null) this.CurrentSlide.Notes = new SlideZoneStructure();
            if (this.CurrentSlide.Notes.Text == null) this.CurrentSlide.Notes.Text = new TextStructure();

            if(new MarkdownBlockBLO().IsImage(markdownBlock))
                this.CurrentSlide.Notes.Text.Text += $"{contentNumber} | <--- [image] {markdownBlockBLO.GetImageInline(markdownBlock)?.Tooltip}...";
            else
            {
                string expressString = (markdownBlock.ToString().Count() < 10)? markdownBlock.ToString() : markdownBlock.ToString().Substring(0, 10);
                this.CurrentSlide.Notes.Text.Text += $"{contentNumber} | <--- {expressString}...";
            }
               

            this.CurrentSlide.Notes.Text.Text += "\r";
        }
        #endregion

    }
}

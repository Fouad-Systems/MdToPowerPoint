using Microsoft.Toolkit.Parsers.Markdown.Blocks;
using PowerPointLibrary.Entities;
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
        PresentationBLO _PresentationBLO;
        SlideZoneStructureBLO _SlideZoneStructureBLO;
        TemplateStructureBLO _TemplateStructureBLO;

        public SlideBLO(PresentationBLO _PresentationBLO)
        {
            this._PresentationBLO = _PresentationBLO;
            _SlideZoneStructureBLO = new SlideZoneStructureBLO();
            _TemplateStructureBLO = new TemplateStructureBLO(_PresentationBLO._PresentationStructure);
        }

        public SlideStructure CurrentSlide
        {
            get
            {
                return _PresentationBLO.CurrentSlide;
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

        public void AddSlide(string Layout)
        {
            // Add Template Zone to Slide
            var TemplateSlide = _TemplateStructureBLO.GetSlide(Layout);

            SlideStructure slideStructure = new SlideStructure();
            _PresentationBLO._PresentationStructure.Slides.Add(slideStructure);

            slideStructure.Name = "Slide" + _PresentationBLO._PresentationStructure.Slides.Count;
            slideStructure.Layout = Layout;
            slideStructure.TemplateSlide = TemplateSlide;
            slideStructure.SlideZones = TemplateSlide.SlideZones.Select(s => s.Clone() as SlideZoneStructure).ToList();


        }


        public void WriteToTitleZone()
        {
            int CurrentZoneOrder = 0;
            if (this.CurrentSlide.CurrentZone != null) CurrentZoneOrder = this.CurrentSlide.CurrentZone.Order;


            this.CurrentSlide.CurrentZone = this.CurrentSlide
                .SlideZones.Where(s => s.Order > CurrentZoneOrder).Where(z => z.ContentTypes.Contains(Entities.Enums.ContentTypes.Title))
                .FirstOrDefault();
        }


        public void WriteToTextZone()
        {
            // Change if we are in zone title or image zone

            if (this.CurrentSlide.IsGenerated == false)
            {
                if (this.CurrentSlide.CurrentZone == null || !this.CurrentSlide.CurrentZone.IsImage())
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
                if (this.CurrentSlide.CurrentZone == null || !this.CurrentSlide.CurrentZone.IsImage())
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

        public void WriteToImageZone()
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

        public void StartWriteToNote()
        {
            _PresentationBLO.CurrentSlide.AddToNotes = true;
        }

        public void EndWriteToNote()
        {
            _PresentationBLO.CurrentSlide.AddToNotes = false;
        }

        public void UseSlide(CommentAction commentAction)
        {
            this.CurrentSlide.UseSlideOrder = commentAction.UseSlideOrder;
        }

        public void AddNotes(MarkdownBlock paragraph)
        {
            if (this.CurrentSlide.Notes == null) this.CurrentSlide.Notes = new SlideZoneStructure();
            if (this.CurrentSlide.Notes.Text == null) this.CurrentSlide.Notes.Text = new TextStructure();

            _SlideZoneStructureBLO.AddMarkdownBlockToSlideZone(this.CurrentSlide.Notes, paragraph);
            this.CurrentSlide.Notes.Text.Text += "\r";
        }



     

        //public void CopySlideZoneToGeneratedSlideZone()
        //{
        //    // Add the created Zone to GeneratedSlideZones
        //    foreach (var item in this.CurrentSlide.SlideZones)
        //    {
        //        if(item.Text != null && !string.IsNullOrEmpty(item.Text.Text))
        //        {

        //        }
        //        if (_SlideZoneStructureBLO.IsHaveData(item))
        //        {
        //            this.CurrentSlide.GeneratedSlideZones.Add(item.Clone() as SlideZoneStructure);
        //        }
        //    }
        //}
    }
}

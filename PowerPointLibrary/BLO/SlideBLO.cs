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

        public SlideBLO(PresentationBLO _PresentationBLO)
        {
            this._PresentationBLO = _PresentationBLO;
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

            var TemplateSlide = _PresentationBLO._TemplateStructure
                .Slides.Where(s => s.Layout == layout).FirstOrDefault();

            if(TemplateSlide == null)
            {
                string msg = $"The layout {layout} doesn't exist";
                throw new PplException(msg);
            }

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
            var TemplateSlide = _PresentationBLO._TemplateStructure.Slides
                 .Where(s => s.Layout == Layout).FirstOrDefault();

            if (TemplateSlide == null)
            {
                string msg = $"The layout {Layout} doesn't exist";
                throw new PowerPointLibrary.Exceptions.PplException(msg);
            }

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
                .SlideZones.Where(s=>s.Order > CurrentZoneOrder).Where(z => z.ContentTypes.Contains(Entities.Enums.ContentTypes.Title))
                .FirstOrDefault();
        }


        public void WriteToTextZone()
        {
            // Change if we are in zone title or image zone

            if(this.CurrentSlide.CurrentZone == null || 
                ! this.CurrentSlide.CurrentZone.ContentTypes.Contains(Entities.Enums.ContentTypes.Text))
            {

                int CurrentZoneOrder = 0;
                if (this.CurrentSlide.CurrentZone != null) CurrentZoneOrder = this.CurrentSlide.CurrentZone.Order;

                this.CurrentSlide.CurrentZone = this.CurrentSlide
              .SlideZones.Where(s => s.Order > CurrentZoneOrder).Where(z => z.ContentTypes.Contains(Entities.Enums.ContentTypes.Text))
              .FirstOrDefault();
            }
           

               
        }

        public void WriteToImageZone()
        {


            int CurrentZoneOrder = 0;
            if (this.CurrentSlide.CurrentZone != null) CurrentZoneOrder = this.CurrentSlide.CurrentZone.Order;

            this.CurrentSlide.CurrentZone = this.CurrentSlide
         .SlideZones.Where(s => s.Order > CurrentZoneOrder).Where(z => z.ContentTypes.Contains(Entities.Enums.ContentTypes.Image))
         .FirstOrDefault();







        }

        public void StartWriteToNote()
        {
            _PresentationBLO.CurrentSlide.AddToNotes = true;
        }

        public void EndWriteToNote()
        {
            _PresentationBLO.CurrentSlide.AddToNotes = false;
        }
    }
}

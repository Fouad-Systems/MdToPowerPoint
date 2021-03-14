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



    public class PresentationStructureBLO
    {
        PresentationStructure _PresentationStructure;

        public PresentationStructureBLO(PresentationStructure presentationStructure)
        {
            _PresentationStructure = presentationStructure;
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
                    slideZone.Order = order++ ;
                }

            }

            return templateStructure;
        }

       

      
    }
}

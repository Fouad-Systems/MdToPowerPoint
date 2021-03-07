using Newtonsoft.Json;
using PowerPointLibrary.Entities;
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
        public void CreateTemplateStructureExemple()
        {

            PresentationStructure templateStructure = new PresentationStructure();

            SlideStructure slide1 = new SlideStructure();
            slide1.Name = "Slide 1";
            slide1.Order = 1;
            slide1.SlideZones.Add(new SlideZoneStructure() { Name = "zone1" });
            slide1.SlideZones.Add(new SlideZoneStructure() { Name = "zone2" });

            SlideStructure slide2 = new SlideStructure();
            slide2.Name = "Slide 2";
            slide2.Order = 1;
            slide2.SlideZones.Add(new SlideZoneStructure() { Name = "zone1" });
            slide2.SlideZones.Add(new SlideZoneStructure() { Name = "zone2" });
            slide2.SlideZones.Add(new SlideZoneStructure() { Name = "zone3" });

            templateStructure.Slides.Add(slide1);
            templateStructure.Slides.Add(slide2);

            File.WriteAllText(@"template.json", JsonConvert.SerializeObject(templateStructure));

            
        }

        public PresentationStructure LoadConfiguration(string templateName)
        {
            string filePath = Environment.CurrentDirectory + "/" + templateName + ".json";
            
            if(!File.Exists(filePath))
            {
                string msg = $"The file {filePath} doesn't exist ";
                throw new PowerPointLibrary.Exceptions.PplException(msg);
            }

            string code = File.ReadAllText(filePath);

            var obj = JsonConvert.DeserializeObject(code, typeof(PresentationStructure));
            PresentationStructure templateStructure = obj as PresentationStructure;
            return templateStructure;
        }
    }
}

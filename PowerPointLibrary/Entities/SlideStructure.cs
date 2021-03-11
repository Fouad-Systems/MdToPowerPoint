using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using PowerPointLibrary.Entities.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.Entities
{




    public class SlideStructure
    {


        public string Name { get; set; }

        public string Layout { get; set; }

        /// <summary>
        /// Is layout changed by Action 
        /// </summary>
        public bool IsLayoutChangedByAction { get; internal set; }


        public int Order { get; set; }
        public List<SlideZoneStructure> SlideZones { get; set; }
      
        public SlideStructure TemplateSlide { set; get; }

        /// <summary>
        /// Curent zone name used to add Data
        /// </summary>
        public SlideZoneStructure CurrentZone { get; internal set; }

        public SlideZoneStructure Notes { set; get; }
        public bool AddToNotes { get; internal set; }

        /// <summary>
        /// Indicate the slide number to be used from the file OutputfileName.slides.pptx
        /// </summary>
        public int UseSlideOrder { get; set; }
       

        public SlideStructure()
        {

            UseSlideOrder = 0;

            this.SlideZones = new List<SlideZoneStructure>();
            Notes = new SlideZoneStructure();
            Notes.Name = "Notes";
          
        }
    }
}

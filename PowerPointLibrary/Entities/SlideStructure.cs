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

        public string Template { get; internal set; }
        public int Order { get; set; }
        public List<SlideZoneStructure> SlideZones { get; set; }
      
        public SlideStructure TemplateSlide { set; get; }


        public SlideStructure()
        {
            this.SlideZones = new List<SlideZoneStructure>();
        }
    }
}

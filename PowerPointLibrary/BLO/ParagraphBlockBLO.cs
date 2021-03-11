using Microsoft.Toolkit.Parsers.Markdown.Blocks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.BLO
{
    public class ParagraphBlockBLO
    {
        public bool IsImage(ParagraphBlock paragraph)
        {
           
           var image =  paragraph.Inlines.First();
            if(image.Type == Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.Image)
                return true;
            else
                return false;
          
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.Entities
{
    public class PplArguments
    {
        public string CurrentDirectoryPath { get; set; }

        public string MdDocumentDirectoryPath { get; set; }
        public string MdDocumentPath { get; set; }

        public string TemplatePath { get; set; }

        public string TemplateConfigurationPath { get; set; }
        public string OutPutPath { get; set; }
        public string OutPutFile { get; set; }

        /// <summary>
        /// Presentation used to find a slides to copy in current presentation
        /// </summary>
        public string UseSlideOutPutFile { get; set; }
    }
}

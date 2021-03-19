using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Toolkit.Parsers.Markdown;
using Microsoft.Toolkit.Parsers.Markdown.Blocks;
using Microsoft.Toolkit.Parsers.Markdown.Inlines;
using PowerPointLibrary.Helper;
using PowerPointLibrary.BLO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using PowerPointLibrary.Exceptions;
using PowerPointLibrary.Entities;

namespace MdToPowerPoint
{
    class Program
    {
        static void Main(string[] args)
        {

          //  PresentationStructureBLO.CreateTemplateStructureExemple();

            PplArguments pplArguments = new PplArgumentsBLO().Read(args);

            //string MdDocumentFileName = "mdData.md";
            //// new TemplateStructureBLO().CreateTemplateStructureExemple();

            //if (args != null && args.Length > 0)
            //{
            //    MdDocumentFileName = args[0];

            //}

            //string FilePath = Environment.CurrentDirectory + "\\" + MdDocumentFileName;
            //if (!File.Exists(FilePath))
            //{
            //    string msg = $"The file ({FilePath}) doesn't exist";
            //    throw new PplException(msg);
            //}


            CreatePresenytation(pplArguments);

            // Clean up the unmanaged PowerPoint COM resources by forcing a  
            // garbage collection as soon as the calling function is off the  
            // stack (at which point these objects are no longer rooted). 
            GC.Collect();
            GC.WaitForPendingFinalizers();
            // GC needs to be called twice in order to get the Finalizers called  
            // - the first time in, it simply makes a list of what is to be  
            // finalized, the second time in, it actually is finalizing. Only  
            // then will the object do its automatic ReleaseComObject. 
            GC.Collect();
            GC.WaitForPendingFinalizers();

        }

        private static void CreatePresenytation(PplArguments pplArguments)
        {


            // Load MarkDown File
            StreamReader sr = new StreamReader(pplArguments.MdDocumentPath);
            string md = sr.ReadToEnd();


            // Parse Markdonw file
            MarkdownDocument mdDocument = new MarkdownDocument();
            mdDocument.Parse(md);


            // Create presentation
            PresentationBLO presentationBLO = new PresentationBLO(pplArguments);
            presentationBLO.Create();

            presentationBLO.CreatePresentation(mdDocument);

            


           // System.Diagnostics.Process.Start(Environment.CurrentDirectory);
            System.Diagnostics.Process.Start(pplArguments.OutPutFile);

        }


    }

    class Presentee
    {
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string Initial { get; set; }
        public string Faculty { get; set; }
        public string Directory { get; set; }
    }
}

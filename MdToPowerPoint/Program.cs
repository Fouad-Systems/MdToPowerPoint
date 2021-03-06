using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Toolkit.Parsers.Markdown;
using Microsoft.Toolkit.Parsers.Markdown.Blocks;
using Microsoft.Toolkit.Parsers.Markdown.Inlines;
using PowerPointLibrary.Helper;
using PowerPointLibrary.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MdToPowerPoint
{
    class Program
    {
        static void Main(string[] args)
        {
            AutoPresentation();
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

        private static void AutoPresentation()
        {


          //  TestHelpers.Main1();

            // Load MarkDown File
            string BaseDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            StreamReader sr = new StreamReader(BaseDir + "\\introduction.md");
            string md = sr.ReadToEnd();


            // Parse
            MarkdownDocument document = new MarkdownDocument();
            document.Parse(md);


            PresentationHelper presentationHelper = new PresentationHelper();
            presentationHelper.Create("template1.potm");


            int SlideIndex = 1;
            bool newSlide = false;
            TitleAndContentSlideHelper slide = null;
            foreach (var element in document.Blocks)
            {
                if (element is HeaderBlock header)
                {

                    slide = new TitleAndContentSlideHelper(presentationHelper, SlideIndex++);

                    string Slide_name = slide.Slide.Name;
                    int c = slide.Slide.Shapes.Count;
                    string name = slide.Slide.Shapes[1].Name;


                    TextRange TitleTextRange = slide.Slide.Shapes[1].TextFrame.TextRange;
                    new TextRangeHelper(TitleTextRange).AddMarkdownBlock(element);

                    //oText.Text = "Bonjour l'informatique";
                    //oText.Words(1, 1).Find("Bonjour").Font.Bold = MsoTriState.msoCTrue;
                    //oText.Words(1, 1).Find("Bonjour").Font.Size = 20;
                    //oText =  oText.Words(1,1).InsertAfter(oText.Words(1, 1));

                }
 
                if (element is ParagraphBlock Paragraph)
                {
                    if(Paragraph.Inlines[0].Type == MarkdownInlineType.Comment)
                    {
                        string comment = Paragraph.Inlines[0].ToString();

                        // Change Slide layout
                        if(comment.StartsWith("<!-- slide : "))
                        {
                            string layout = comment.Replace("<!-- slide : ", "");
                            layout = layout.Replace("-->", "");
                            layout = layout.Trim();
                            // slide = new TitleAndContentSlideHelper(presentationHelper, SlideIndex++);
                            slide.ChangeLayout(layout);
                        }

                        // Change zone
                        if (comment.StartsWith("<!-- zone : "))
                        {
                            string ShapesName = comment.Replace("<!-- zone : ", "");
                            ShapesName = ShapesName.Replace("-->", "");
                            ShapesName = ShapesName.Trim();
                            // slide = new TitleAndContentSlideHelper(presentationHelper, SlideIndex++);
                            slide.CurrentShapesName = ShapesName;
                        }

                    }

                    TextRange TitleTextRange = slide.Slide.Shapes[2].TextFrame.TextRange;
                    if (!string.IsNullOrEmpty(slide.CurrentShapesName))
                    {
                        string Slide_name = slide.Slide.Name;
                        int c = slide.Slide.Shapes.Count;
                        string name = slide.Slide.Shapes[1].Name;
                         name  = slide.Slide.Shapes[2].Name;
                        name = slide.Slide.Shapes[3].Name;
                        TitleTextRange = slide.Slide.Shapes["Content Placeholder 6"].TextFrame.TextRange;
                    }
                   
                    new TextRangeHelper(TitleTextRange).AddMarkdownBlock(element);
                   
                    
                    
                }

            }


            string fileName = BaseDir + "\\output.pptx";
            presentationHelper.SaveAs(fileName);
            presentationHelper.Close();

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

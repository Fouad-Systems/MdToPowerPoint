using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Toolkit.Parsers.Markdown.Blocks;
using Microsoft.Toolkit.Parsers.Markdown.Inlines;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.Helper
{
    public class TextRangeHelper
    {
        protected TextRange TextRange;

        public TextRangeHelper(TextRange TextRange)
        {
            this.TextRange = TextRange;
        }

        //[Obsolete]
        //public void AddMarkdownBlock(MarkdownBlock markdownBlock)
        //{

        //    switch (markdownBlock.Type)
        //    {
        //        case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.Root:
        //            break;
        //        case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.Paragraph:
        //            Microsoft.Toolkit.Parsers.Markdown.Blocks.ParagraphBlock ParagraphBlock = markdownBlock as ParagraphBlock;
        //            this.AddInLindes(ParagraphBlock.Inlines);
        //            break;
        //        case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.Quote:
        //            break;
        //        case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.Code:
        //            break;
        //        case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.Header:

        //            HeaderBlock headerBlock = markdownBlock as HeaderBlock;
        //            this.AddInLindes(headerBlock.Inlines);

        //            break;
        //        case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.List:
        //            break;
        //        case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.ListItemBuilder:
        //            break;
        //        case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.HorizontalRule:
        //            break;
        //        case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.Table:
        //            break;
        //        case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.LinkReference:
        //            break;
        //        case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.YamlHeader:
        //            break;
        //        default:
        //            break;
        //    }
 
        //}

       

       


        //protected void AddInLindes(IList<MarkdownInline> MarkdownInlines)
        //{

           

        //    foreach (var markdownInline in MarkdownInlines)
        //    {
        //        switch (markdownInline.Type)
        //        {
        //            case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.Comment:
        //                break;
        //            case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.TextRun:
        //                TextRunInline textRunInline = markdownInline as TextRunInline;
        //                string text = textRunInline.Text;
        //                this.TextRange.InsertAfter(text);
        //                break;
        //            case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.Bold:

                      
        //                int Start = this.TextRange.Text.Count() + 1;
                       
        //                Microsoft.Toolkit.Parsers.Markdown.Inlines.BoldTextInline boldTextInline = markdownInline as BoldTextInline;
        //                string text_blod = (boldTextInline.Inlines[0] as TextRunInline).Text;
        //                int Length = text_blod.Count();
        //                this.TextRange.InsertAfter(text_blod);

        //                TextElements.Add(new TextElement(Start, Length, TextElement.TextStyles.Blod));

                       



        //                break;
        //            case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.Italic:
        //                int StartItalic = this.TextRange.Text.Count() + 1;
        //                Microsoft.Toolkit.Parsers.Markdown.Inlines.ItalicTextInline italicTextInline = markdownInline as ItalicTextInline;
        //                string text_italic = (italicTextInline.Inlines[0] as TextRunInline).Text;
        //                int Lengthitalic = text_italic.Count();
        //                this.TextRange.InsertAfter(text_italic);
        //                TextElements.Add(new TextElement(StartItalic, Lengthitalic, TextElement.TextStyles.Italic));
        //                break;
        //            case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.MarkdownLink:
        //                break;
        //            case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.RawHyperlink:
        //                break;
        //            case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.RawSubreddit:
        //                break;
        //            case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.Strikethrough:
        //                break;
        //            case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.Superscript:
        //                break;
        //            case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.Subscript:
        //                break;
        //            case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.Code:
        //                break;
        //            case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.Image:
        //                break;
        //            case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.Emoji:
        //                break;
        //            case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.LinkReference:
        //                break;
        //            default:
        //                break;
        //        }
        //    }


        //    foreach (var textElement in TextElements)
        //    {

        //        TextRange textRange = this.TextRange.Characters(textElement.Start, textElement.Length);
        //        switch (textElement.TextStyle)
        //        {
        //            case TextElement.TextStyles.Blod:
        //                textRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoCTrue;
        //                break;
        //            case TextElement.TextStyles.Italic:
        //                textRange.Font.Italic = Microsoft.Office.Core.MsoTriState.msoCTrue;
        //                break;
        //            default:
        //                break;
        //        }
               
        //    }
            
        //}
    }
}

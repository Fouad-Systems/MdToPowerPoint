using Microsoft.Toolkit.Parsers.Markdown.Blocks;
using Microsoft.Toolkit.Parsers.Markdown.Inlines;
using PowerPointLibrary.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.BLO
{
    public class TextStructureBLO
    {

       

        public TextStructure CreateAndAddFromMarkdownBlock(TextStructure textStructure, MarkdownBlock markdownBlock)
        {
          
            switch (markdownBlock.Type)
            {
                case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.Root:
                    break;
                case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.Paragraph:
                    Microsoft.Toolkit.Parsers.Markdown.Blocks.ParagraphBlock ParagraphBlock = markdownBlock as ParagraphBlock;
                    this.AddInLindesToTextStructure(textStructure, ParagraphBlock.Inlines);
                    break;
                case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.Quote:
                    break;
                case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.Code:
                    break;
                case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.Header:

                    HeaderBlock headerBlock = markdownBlock as HeaderBlock;
                    this.AddInLindesToTextStructure(textStructure, headerBlock.Inlines);

                    break;
                case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.List:
                    break;
                case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.ListItemBuilder:
                    break;
                case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.HorizontalRule:
                    break;
                case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.Table:
                    break;
                case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.LinkReference:
                    break;
                case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.YamlHeader:
                    break;
                default:
                    break;
            }

            return textStructure;
        }


        protected void AddInLindesToTextStructure(TextStructure textStructure, IList<MarkdownInline> MarkdownInlines)
        {

            foreach (var markdownInline in MarkdownInlines)
            {
                switch (markdownInline.Type)
                {
                    case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.Comment:
                        break;
                    case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.TextRun:
                        TextRunInline textRunInline = markdownInline as TextRunInline;
                        string text = textRunInline.Text;
                        textStructure.Text += text;
                        break;
                    case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.Bold:
                        int Start = textStructure.Text.Count() + 1;
                        Microsoft.Toolkit.Parsers.Markdown.Inlines.BoldTextInline boldTextInline = markdownInline as BoldTextInline;
                        string text_blod = (boldTextInline.Inlines[0] as TextRunInline).Text;
                        int Length = text_blod.Count();
                        textStructure.Text +=  text_blod;
                        textStructure.TextElementStyles.Add(new TextElementStyle(Start, Length, TextElementStyle.TextStyles.Blod));
                        break;
                    case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.Italic:
                        int StartItalic = textStructure.Text.Count() + 1;
                        Microsoft.Toolkit.Parsers.Markdown.Inlines.ItalicTextInline italicTextInline = markdownInline as ItalicTextInline;
                        string text_italic = (italicTextInline.Inlines[0] as TextRunInline).Text;
                        int Lengthitalic = text_italic.Count();
                        textStructure.Text += text_italic;
                        textStructure.TextElementStyles.Add(new TextElementStyle(StartItalic, Lengthitalic, TextElementStyle.TextStyles.Italic));
                        break;
                    case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.MarkdownLink:
                        break;
                    case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.RawHyperlink:
                        break;
                    case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.RawSubreddit:
                        break;
                    case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.Strikethrough:
                        break;
                    case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.Superscript:
                        break;
                    case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.Subscript:
                        break;
                    case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.Code:
                        break;
                    case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.Image:
                        break;
                    case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.Emoji:
                        break;
                    case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.LinkReference:
                        break;
                    default:
                        break;
                }
            }
        }
    }
}

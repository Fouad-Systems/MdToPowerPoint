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
    public class SlideZoneStructureBLO
    {

        public void AddMarkdownBlockToSlideZone(SlideZoneStructure SlideZone, MarkdownBlock markdownBlock)
        {

            switch (markdownBlock.Type)
            {
                case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.Root:
                    break;
                case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.Paragraph:
                    Microsoft.Toolkit.Parsers.Markdown.Blocks.ParagraphBlock ParagraphBlock = markdownBlock as ParagraphBlock;
                    this.AddInLindesToSlideZone(SlideZone, ParagraphBlock.Inlines);
                    break;
                case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.Quote:
                    break;
                case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.Code:
                    break;
                case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.Header:

                    HeaderBlock headerBlock = markdownBlock as HeaderBlock;
                    this.AddInLindesToSlideZone(SlideZone, headerBlock.Inlines);

                    break;
                case Microsoft.Toolkit.Parsers.Markdown.MarkdownBlockType.List:
                    Microsoft.Toolkit.Parsers.Markdown.Blocks.ListBlock listBlock = markdownBlock as ListBlock;
                    this.AddListBlockToTextStructure(SlideZone, listBlock);
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

        }

        private void AddListBlockToTextStructure(SlideZoneStructure SlideZone, ListBlock listBlock)
        {
            for (int i = 0; i < listBlock.Items.Count; i++)
            {
                foreach (MarkdownBlock markdownBlock in listBlock.Items[i].Blocks)
                {
                    int Start = SlideZone.Text.Text.Count() + 1;
                    this.AddMarkdownBlockToSlideZone(SlideZone, markdownBlock);
                    int Length = SlideZone.Text.Text.Count() - Start;
                    SlideZone.Text.TextElementStyles
                        .Add(new TextElementStyle(Start, Length, TextElementStyle.TextStyles.Bullet) { ListStyle = listBlock.Style });

                }

                // ne pas ajouter le retour à la ligne pour  le dernier élément.
                if (i < listBlock.Items.Count - 1)
                    SlideZone.Text.Text += "\n";
            }


        }

        protected void AddInLindesToSlideZone(SlideZoneStructure SlideZone, IList<MarkdownInline> MarkdownInlines)
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
                        SlideZone.Text.Text += text;
                        break;
                    case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.Bold:
                        int Start = SlideZone.Text.Text.Count() + 1;
                        Microsoft.Toolkit.Parsers.Markdown.Inlines.BoldTextInline boldTextInline = markdownInline as BoldTextInline;
                        string text_blod = (boldTextInline.Inlines[0] as TextRunInline).Text;
                        int Length = text_blod.Count();
                        SlideZone.Text.Text += text_blod;
                        SlideZone.Text.TextElementStyles.Add(new TextElementStyle(Start, Length, TextElementStyle.TextStyles.Blod));
                        break;
                    case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.Italic:
                        int StartItalic = SlideZone.Text.Text.Count() + 1;
                        Microsoft.Toolkit.Parsers.Markdown.Inlines.ItalicTextInline italicTextInline = markdownInline as ItalicTextInline;
                        string text_italic = (italicTextInline.Inlines[0] as TextRunInline).Text;
                        int Lengthitalic = text_italic.Count();
                        SlideZone.Text.Text += text_italic;
                        SlideZone.Text.TextElementStyles.Add(new TextElementStyle(StartItalic, Lengthitalic, TextElementStyle.TextStyles.Italic));
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
                        Microsoft.Toolkit.Parsers.Markdown.Inlines.ImageInline imageInline = markdownInline as ImageInline;
                        if (SlideZone.Images == null) SlideZone.Images = new List<ImageStructure>();
                        ImageStructure imageStructure = new ImageStructure();
                        SlideZone.Images.Add(imageStructure);


                        imageStructure.ImageHeight = imageInline.ImageHeight;
                        imageStructure.ImageWidth = imageInline.ImageWidth;
                        imageStructure.Url = imageInline.Url;
                        imageStructure.Tooltip = imageInline.Tooltip;
                        //string text = textRunInline.Text;
                        //textStructure.Text += text;
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

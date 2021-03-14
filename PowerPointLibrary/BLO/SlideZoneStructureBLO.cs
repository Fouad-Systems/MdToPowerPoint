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

                    this.TrimFirstInlines(headerBlock);



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

        private void TrimFirstInlines(HeaderBlock headerBlock)
        {
            var first = headerBlock.Inlines.First();

            switch (first.Type)
            {
                case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.TextRun:
                    (first as TextRunInline).Text = (first as TextRunInline).Text.Remove(0, 1);
                    break;
                case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.Bold:
                    ((first as BoldTextInline).Inlines[0] as TextRunInline).Text = ((first as BoldTextInline).Inlines[0] as TextRunInline).Text.Remove(0, 1);
                    break;
                case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.Italic:
                    ((first as ItalicTextInline).Inlines[0] as TextRunInline).Text = ((first as ItalicTextInline).Inlines[0] as TextRunInline).Text.Remove(0, 1);
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

                        // in head , a space is auto-added, we must delete it
                         string text = textRunInline.Text;

                        if (SlideZone.Text == null) SlideZone.Text = new TextStructure();
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
                        int start_ItalicTextInline = SlideZone.Text.Text.Count() + 1;
                        Microsoft.Toolkit.Parsers.Markdown.Inlines.ItalicTextInline o_ItalicTextInline = markdownInline as ItalicTextInline;
                        string text_ItalicTextInline = (o_ItalicTextInline.Inlines[0] as TextRunInline).Text;
                        int Length_ItalicTextInline = text_ItalicTextInline.Count();
                        SlideZone.Text.Text += text_ItalicTextInline;
                        SlideZone.Text.TextElementStyles.Add(new TextElementStyle(start_ItalicTextInline, Length_ItalicTextInline, TextElementStyle.TextStyles.Italic));
                        break;
                    case Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.MarkdownLink:
                       
                        Microsoft.Toolkit.Parsers.Markdown.Inlines.MarkdownLinkInline o_MarkdownLinkInline = markdownInline as MarkdownLinkInline;

                        this.AddInLindesToSlideZone(SlideZone, o_MarkdownLinkInline.Inlines);
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

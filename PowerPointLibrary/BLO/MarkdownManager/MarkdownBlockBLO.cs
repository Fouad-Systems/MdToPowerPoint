using Microsoft.Toolkit.Parsers.Markdown;
using Microsoft.Toolkit.Parsers.Markdown.Blocks;
using Microsoft.Toolkit.Parsers.Markdown.Inlines;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.BLO
{
    public class MarkdownBlockBLO
    {
        public bool IsImage(MarkdownBlock markdownBlock)
        {
            ParagraphBlock paragraphBlock = markdownBlock as ParagraphBlock;
            if (paragraphBlock == null) return false;
            return this.IsImage(paragraphBlock);
        }

        public bool IsImage(ParagraphBlock paragraph)
        {

            var image = paragraph.Inlines.First();
            if (image.Type == Microsoft.Toolkit.Parsers.Markdown.MarkdownInlineType.Image)
                return true;
            else
                return false;

        }

        public ImageInline GetImageInline(MarkdownBlock paragraph)
        {

            ParagraphBlock paragraphBlock = paragraph as ParagraphBlock;
            if (paragraphBlock == null) return null;

            var markdownInline = paragraphBlock.Inlines.First();
            ImageInline imageInline = markdownInline as ImageInline;
            return imageInline;

        }

        public bool isAction(MarkdownBlock element)
        {
            ParagraphBlock paragraphBlock = element as ParagraphBlock;
            if (paragraphBlock == null) return false;

            CommentActionBLO commentActionBLO = new CommentActionBLO();
            if (paragraphBlock.Inlines[0].Type == MarkdownInlineType.Comment
                && commentActionBLO.IsAction(paragraphBlock.Inlines[0].ToString()))
            {
                return true;
            }
            return false;
        }

        public string GetComment(MarkdownBlock element)
        {
            ParagraphBlock paragraphBlock = element as ParagraphBlock;
            if (paragraphBlock == null) return null ;
            return paragraphBlock.Inlines[0].ToString();
        }
    }
}

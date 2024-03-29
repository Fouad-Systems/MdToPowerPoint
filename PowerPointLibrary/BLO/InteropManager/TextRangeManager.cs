﻿using Microsoft.Office.Interop.PowerPoint;
using PowerPointLibrary.Entities;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.Manager
{
    // Add Texte to TexteRange
    public class TextRangeManager
    {
        public void AddTextStructure(TextRange textRange, TextStructure textStructure)
        {
            textRange.Text = textStructure.Text;
           // textRange.ParagraphFormat.Alignment = textStructure.ParagraphAlignment;

            foreach (var textElement in textStructure.TextElementStyles)
            {

                var source = textRange.Text;

                TextRange textRangePart = textRange.Characters(textElement.Start, textElement.Length);
                //  textRangePart.Font.Bold = Microsoft.Office.Core.MsoTriState.msoCTrue;

                var text = textRangePart.Text;

                if (textElement.IsBlod)
                    textRangePart.Font.Bold = Microsoft.Office.Core.MsoTriState.msoCTrue;
                if (textElement.IsItalic)
                    textRangePart.Font.Italic = Microsoft.Office.Core.MsoTriState.msoCTrue;

                if (!string.IsNullOrEmpty(textElement.FontColor))
                {
                   
                    textRangePart.Font.Bold = Microsoft.Office.Core.MsoTriState.msoCTrue;
                    var fromHTML = ColorTranslator.FromHtml(textElement.FontColor);
                    var fromHTML2 = ColorTranslator.ToHtml(fromHTML);
                    var Color1 = ColorTranslator.FromHtml(fromHTML2);
                    // var Color1 = ColorTranslator.FromHtml("#068dbd");
                    textRangePart.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(Color1);
                    // textRangePart.Font.Color.RGB = Color.Red.ToArgb();
                }

                if (textElement.IsBullet)
                {
                    switch (textElement.ListStyle)
                    {
                        case Microsoft.Toolkit.Parsers.Markdown.ListStyle.Bulleted:
                            textRangePart.ParagraphFormat.Bullet.Type = PpBulletType.ppBulletUnnumbered;
                            textRangePart.ParagraphFormat.Bullet.Character = 9632;


                            textRangePart.ParagraphFormat.Bullet.Visible = Microsoft.Office.Core.MsoTriState.msoCTrue;
                            break;
                        case Microsoft.Toolkit.Parsers.Markdown.ListStyle.Numbered:
                            textRangePart.ParagraphFormat.Bullet.Type = PpBulletType.ppBulletNumbered;
                            textRangePart.ParagraphFormat.Bullet.Visible = Microsoft.Office.Core.MsoTriState.msoCTrue;
                            // textRangePart.ParagraphFormat.Bullet.Style = PpNumberedBulletStyle.ppBulletCircleNumDBPlain;
                            break;
                        default:
                            break;
                    }
                }

            }
        }
    }
}

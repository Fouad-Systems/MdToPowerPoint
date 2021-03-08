using Microsoft.Office.Interop.PowerPoint;
using PowerPointLibrary.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.Manager
{
    public class TextRangeManager
    {
        public void AddTextStructure(TextRange textRange, TextStructure textStructure)
        {
            textRange.Text = textStructure.Text;

            foreach (var textElement in textStructure.TextElementStyles)
            {

                TextRange textRangePart = textRange.Characters(textElement.Start, textElement.Length);
                switch (textElement.TextStyle)
                {
                    case TextElementStyle.TextStyles.Blod:
                        textRangePart.Font.Bold = Microsoft.Office.Core.MsoTriState.msoCTrue;
                        break;
                    case TextElementStyle.TextStyles.Italic:
                        textRangePart.Font.Italic = Microsoft.Office.Core.MsoTriState.msoCTrue;
                        break;
                    case TextElementStyle.TextStyles.Bullet:
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
                        break;
                    default:
                        break;
                }

            }
        }
    }
}

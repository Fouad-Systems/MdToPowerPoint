using ColorCode.Common;
using ColorCode.Styling;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.BLO
{
    public class CodeBLO
    {

        public const string Blue = "#aaa";
        public const string DarkCyan = "#FF008B8B";
        public const string DarkOliveGreen = "#FF556B2F";
        public const string OliveDrab = "#FF6B8E23";
        public const string Magenta = "FFFF00FF";
        public const string MediumTurqoise = "FF48D1CC";
        public const string Red = "#FFFF0000";
        public const string Purple = "#FF800080";
        public const string Navy = "#FF000080";
        public const string OrangeRed = "#FFFF4500";
        public const string Teal = "#FF008080";
        public const string PowderBlue = "#FFB0E0E6";
        public const string Green = "#FF008000";
        public const string Yellow = "#FFFFFF00";
        public const string DullRed = "#FFA31515";
        public const string Black = "#FF000000";
        public const string White = "#FFFFFFFF";
        public const string Gray = "#FF808080";



        public  StyleDictionary GetDefaultCodeStyle()
        {
            return new StyleDictionary
                {
                    new Style(ScopeName.PlainText)
                    {
                        Foreground = Black,
                        Background = White,
                        ReferenceName = "plainText"
                    },
                    new Style(ScopeName.HtmlServerSideScript)
                    {
                        Background = Yellow,
                        ReferenceName = "htmlServerSideScript"
                    },
                    new Style(ScopeName.HtmlComment)
                    {
                        Foreground = Green,
                        ReferenceName = "htmlComment"
                    },
                    new Style(ScopeName.HtmlTagDelimiter)
                    {
                        Foreground = Blue,
                        ReferenceName = "htmlTagDelimiter"
                    },
                    new Style(ScopeName.HtmlElementName)
                    {
                        Foreground = DullRed,
                        ReferenceName = "htmlElementName"
                    },
                    new Style(ScopeName.HtmlAttributeName)
                    {
                        Foreground = Red,
                        ReferenceName = "htmlAttributeName"
                    },
                    new Style(ScopeName.HtmlAttributeValue)
                    {
                        Foreground = Blue,
                        ReferenceName = "htmlAttributeValue"
                    },
                    new Style(ScopeName.HtmlOperator)
                    {
                        Foreground = Blue,
                        ReferenceName = "htmlOperator"
                    },
                    new Style(ScopeName.Comment)
                    {
                        Foreground = Green,
                        ReferenceName = "comment"
                    },
                    new Style(ScopeName.XmlDocTag)
                    {
                        Foreground = Gray,
                        ReferenceName = "xmlDocTag"
                    },
                    new Style(ScopeName.XmlDocComment)
                    {
                        Foreground = Green,
                        ReferenceName = "xmlDocComment"
                    },
                    new Style(ScopeName.String)
                    {
                        Foreground = DullRed,
                        ReferenceName = "string"
                    },
                    new Style(ScopeName.StringCSharpVerbatim)
                    {
                        Foreground = DullRed,
                        ReferenceName = "stringCSharpVerbatim"
                    },
                    new Style(ScopeName.Keyword)
                    {
                        Foreground = Blue,
                        ReferenceName = "keyword"
                    },
                    new Style(ScopeName.PreprocessorKeyword)
                    {
                        Foreground = Blue,
                        ReferenceName = "preprocessorKeyword"
                    },
                    new Style(ScopeName.HtmlEntity)
                    {
                        Foreground = Red,
                        ReferenceName = "htmlEntity"
                    },
                    new Style(ScopeName.XmlAttribute)
                    {
                        Foreground = Red,
                        ReferenceName = "xmlAttribute"
                    },
                    new Style(ScopeName.XmlAttributeQuotes)
                    {
                        Foreground = Black,
                        ReferenceName = "xmlAttributeQuotes"
                    },
                    new Style(ScopeName.XmlAttributeValue)
                    {
                        Foreground = Blue,
                        ReferenceName = "xmlAttributeValue"
                    },
                    new Style(ScopeName.XmlCDataSection)
                    {
                        Foreground = Gray,
                        ReferenceName = "xmlCDataSection"
                    },
                    new Style(ScopeName.XmlComment)
                    {
                        Foreground = Green,
                        ReferenceName = "xmlComment"
                    },
                    new Style(ScopeName.XmlDelimiter)
                    {
                        Foreground = Blue,
                        ReferenceName = "xmlDelimiter"
                    },
                    new Style(ScopeName.XmlName)
                    {
                        Foreground = DullRed,
                        ReferenceName = "xmlName"
                    },
                    new Style(ScopeName.ClassName)
                    {
                        Foreground = MediumTurqoise,
                        ReferenceName = "className"
                    },
                    new Style(ScopeName.CssSelector)
                    {
                        Foreground = DullRed,
                        ReferenceName = "cssSelector"
                    },
                    new Style(ScopeName.CssPropertyName)
                    {
                        Foreground = Red,
                        ReferenceName = "cssPropertyName"
                    },
                    new Style(ScopeName.CssPropertyValue)
                    {
                        Foreground = Blue,
                        ReferenceName = "cssPropertyValue"
                    },
                    new Style(ScopeName.SqlSystemFunction)
                    {
                        Foreground = Magenta,
                        ReferenceName = "sqlSystemFunction"
                    },
                    new Style(ScopeName.PowerShellAttribute)
                    {
                        Foreground = PowderBlue,
                        ReferenceName = "powershellAttribute"
                    },
                    new Style(ScopeName.PowerShellOperator)
                    {
                        Foreground = Gray,
                        ReferenceName = "powershellOperator"
                    },
                    new Style(ScopeName.PowerShellType)
                    {
                        Foreground = Teal,
                        ReferenceName = "powershellType"
                    },
                    new Style(ScopeName.PowerShellVariable)
                    {
                        Foreground = OrangeRed,
                        ReferenceName = "powershellVariable"
                    },

                    new Style(ScopeName.Type)
                    {
                        Foreground = Teal,
                        ReferenceName = "type"
                    },
                    new Style(ScopeName.TypeVariable)
                    {
                        Foreground = Teal,
                        Italic = true,
                        ReferenceName = "typeVariable"
                    },
                    new Style(ScopeName.NameSpace)
                    {
                        Foreground = Navy,
                        ReferenceName = "namespace"
                    },
                    new Style(ScopeName.Constructor)
                    {
                        Foreground = Purple,
                        ReferenceName = "constructor"
                    },
                    new Style(ScopeName.Predefined)
                    {
                        Foreground = Navy,
                        ReferenceName = "predefined"
                    },
                    new Style(ScopeName.PseudoKeyword)
                    {
                        Foreground = Navy,
                        ReferenceName = "pseudoKeyword"
                    },
                    new Style(ScopeName.StringEscape)
                    {
                        Foreground = Gray,
                        ReferenceName = "stringEscape"
                    },
                    new Style(ScopeName.ControlKeyword)
                    {
                        Foreground = Blue,
                        ReferenceName = "controlKeyword"
                    },
                    new Style(ScopeName.Number)
                    {
                        ReferenceName = "number"
                    },
                    new Style(ScopeName.Operator)
                    {
                        ReferenceName = "operator"
                    },
                    new Style(ScopeName.Delimiter)
                    {
                        ReferenceName = "delimiter"
                    },

                    new Style(ScopeName.MarkdownHeader)
                    {
                        Foreground = Blue,
                        Bold = true,
                        ReferenceName = "markdownHeader"
                    },
                    new Style(ScopeName.MarkdownCode)
                    {
                        Foreground = Teal,
                        ReferenceName = "markdownCode"
                    },
                    new Style(ScopeName.MarkdownListItem)
                    {
                        Bold = true,
                        ReferenceName = "markdownListItem"
                    },
                    new Style(ScopeName.MarkdownEmph)
                    {
                        Italic = true,
                        ReferenceName = "italic"
                    },
                    new Style(ScopeName.MarkdownBold)
                    {
                        Bold = true,
                        ReferenceName = "bold"
                    },

                    new Style(ScopeName.BuiltinFunction)
                    {
                        Foreground = OliveDrab,
                        Bold = true,
                        ReferenceName = "builtinFunction"
                    },
                    new Style(ScopeName.BuiltinValue)
                    {
                        Foreground = DarkOliveGreen,
                        Bold = true,
                        ReferenceName = "builtinValue"
                    },
                    new Style(ScopeName.Attribute)
                    {
                        Foreground = DarkCyan,
                        Italic = true,
                        ReferenceName = "attribute"
                    },
                    new Style(ScopeName.SpecialCharacter)
                    {
                        ReferenceName = "specialChar"
                    },
                };
        }
    }
}

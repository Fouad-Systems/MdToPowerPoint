// Copyright (c) Microsoft Corporation.  All rights reserved.

using System.Collections.Generic;
using System.IO;
using ColorCode.Common;
using ColorCode.Parsing;
using System.Text;
using ColorCode.HTML.Common;
using ColorCode.Styling;
using System.Net;
using PowerPointLibrary.Entities;
using Microsoft.Toolkit.Parsers.Markdown.Blocks;

namespace ColorCode
{
    /// <summary>
    /// 
    /// </summary>
    public class TexteStructureCodeColorizer : CodeColorizerBase
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Style">The Custom styles to Apply to the formatted Code.</param>
        /// <param name="languageParser">The language parser that the <see cref="TexteStructureCodeColorizer"/> instance will use for its lifetime.</param>
        public TexteStructureCodeColorizer(TextStructure TextStructure, StyleDictionary Style = null, ILanguageParser languageParser = null) : base(Style, languageParser)
        {
            this.TextStructure = TextStructure;
        }

        private TextWriter Writer { get; set; }
        public TextStructure TextStructure { get; set; }



        /// <summary>
        /// 
        /// </summary>
        /// <param name="sourceCode">The source code to colorize.</param>
        /// <param name="language">The language to use to colorize the source code.</param>
        /// <returns>Colorised HTML Markup.</returns>
        public string SetCodeBlock(CodeBlock codeBlock )
        {

            ILanguage language = Languages.FindById(codeBlock.CodeLanguage);

            var buffer = new StringBuilder(codeBlock.Text.Length * 2);

            using (TextWriter writer = new StringWriter(buffer))
            {
                Writer = writer;
                WriteHeader(language);

                languageParser.Parse(codeBlock.Text, language, (parsedSourceCode, captures) => Write(parsedSourceCode, captures));

                WriteFooter(language);

                writer.Flush();
            }

            return buffer.ToString();
        }

        protected override void Write(string parsedSourceCode, IList<Scope> scopes)
        {
            var styleInsertions = new List<TextInsertion>();

            foreach (Scope scope in scopes)
                GetStyleInsertionsForCapturedStyle(scope, styleInsertions);

            styleInsertions.SortStable((x, y) => x.Index.CompareTo(y.Index));

            int offset = 0;

            this.TextStructure.Text += parsedSourceCode;

            foreach (TextInsertion styleInsertion in styleInsertions)
            {
                Style style = null;


                var text = parsedSourceCode.Substring(offset, styleInsertion.Index - offset);

                //var text = parsedSourceCode.Substring(styleInsertion.Index, styleInsertion.Scope.Length);
              
                offset = styleInsertion.Scope.Length;
                this.AddTextElementStyle(styleInsertion);

            }




        }

        private void AddTextElementStyle(TextInsertion TextInsertion)
        {
            var  style = GetStyleForScope(TextInsertion.Scope);

            if (style != null)
            {
                // les caractère sont enlever '\n'
                int Start = TextStructure.Text.Replace("\n","").Length + 1;
                int Length = text.Length;
                this.TextStructure.Text += text;

                TextElementStyle textElementStyle = new TextElementStyle(Start, Length);
                textElementStyle.IsBlod = style.Bold;
                textElementStyle.IsItalic = style.Italic;
                textElementStyle.FontColor = style.Foreground;


                this.TextStructure.TextElementStyles.Add(textElementStyle);

            }
        }

        private void WriteFooter(ILanguage language)
        {
            Writer.WriteLine();
            WriteHeaderPreEnd();
            WriteHeaderDivEnd();
        }

        private void WriteHeader(ILanguage language)
        {
            WriteHeaderDivStart();
            WriteHeaderPreStart();
            Writer.WriteLine();
        }

        private void GetStyleInsertionsForCapturedStyle(Scope scope, ICollection<TextInsertion> styleInsertions)
        {
            styleInsertions.Add(new TextInsertion
            {
                Index = scope.Index,
                Scope = scope
            });

            foreach (Scope childScope in scope.Children)
                GetStyleInsertionsForCapturedStyle(childScope, styleInsertions);

            //styleInsertions.Add(new TextInsertion
            //{
            //    Index = scope.Index + scope.Length,
            //    Text = "</span>"
            //});
        }

        private Style GetStyleForScope(Scope scope)
        {
            Style rStyle = new Style("value");
            rStyle.Foreground = string.Empty;
            rStyle.Background = string.Empty;
            rStyle.Italic = false;
            rStyle.Bold = false;

            if (Styles.Contains(scope.Name))
            {
                Style style = Styles[scope.Name];

                rStyle.Foreground = style.Foreground;
                rStyle.Background = style.Background;
                rStyle.Italic = style.Italic;
                rStyle.Bold = style.Bold;
            }

            return rStyle;
        }

        private void WriteHeaderDivEnd()
        {
            WriteElementEnd("div");
        }

        private void WriteElementEnd(string elementName)
        {
            Writer.Write("</{0}>", elementName);
        }

        private void WriteHeaderPreEnd()
        {
            WriteElementEnd("pre");
        }

        private void WriteHeaderPreStart()
        {
            WriteElementStart("pre");
        }

        private void WriteHeaderDivStart()
        {
            string foreground = string.Empty;
            string background = string.Empty;

            if (Styles.Contains(ScopeName.PlainText))
            {
                Style plainTextStyle = Styles[ScopeName.PlainText];

                foreground = plainTextStyle.Foreground;
                background = plainTextStyle.Background;
            }

            WriteElementStart("div", foreground, background);
        }

        private void WriteElementStart(string elementName, string foreground = null, string background = null, bool italic = false, bool bold = false)
        {
            Writer.Write("<{0}", elementName);

            if (!string.IsNullOrWhiteSpace(foreground) || !string.IsNullOrWhiteSpace(background) || italic || bold)
            {
                Writer.Write(" style=\"");

                if (!string.IsNullOrWhiteSpace(foreground))
                    Writer.Write("color:{0};", foreground.ToHtmlColor());

                if (!string.IsNullOrWhiteSpace(background))
                    Writer.Write("background-color:{0};", background.ToHtmlColor());

                if (italic)
                    Writer.Write("font-style: italic;");

                if (bold)
                    Writer.Write("font-weight: bold;");

                Writer.Write("\"");
            }

            Writer.Write(">");
        }
    }
}
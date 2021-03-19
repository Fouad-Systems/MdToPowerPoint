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
using System.Drawing;

namespace ColorCode
{
    /// <summary>
    /// 
    /// </summary>
    public class TexteStructureCodeColorizer : CodeColorizerBase
    {
        StyleDictionary _StyleDictionary;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Style">The Custom styles to Apply to the formatted Code.</param>
        /// <param name="languageParser">The language parser that the <see cref="TexteStructureCodeColorizer"/> instance will use for its lifetime.</param>
        public TexteStructureCodeColorizer(TextStructure TextStructure, StyleDictionary Style = null, ILanguageParser languageParser = null) : base(Style, languageParser)
        {
            //this._StyleDictionary = Style;
            //if (Style == null)
            //    this._StyleDictionary = new StyleDictionary();
            
            this.TextStructure = TextStructure;
        }

        public TextStructure TextStructure { get; set; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sourceCode">The source code to colorize.</param>
        /// <param name="language">The language to use to colorize the source code.</param>
        /// <returns>Colorised HTML Markup.</returns>
        public void SetCodeBlock(CodeBlock codeBlock)
        {

            ILanguage language = Languages.FindById(codeBlock.CodeLanguage);

            languageParser.Parse(codeBlock.Text, language, (parsedSourceCode, captures) => Write(parsedSourceCode, captures));

        }

        protected override void Write(string parsedSourceCode, IList<Scope> scopes)
        {
            var styleInsertions = new List<TextInsertion>();

            foreach (Scope scope in scopes)
                GetStyleInsertionsForCapturedStyle(scope, styleInsertions);

            styleInsertions.SortStable((x, y) => x.Index.CompareTo(y.Index));

            int offset = this.TextStructure.Text.Length;

            this.TextStructure.Text += parsedSourceCode.Replace("\n", "");

            foreach (TextInsertion styleInsertion in styleInsertions)
            {
                Style style = null;


                // var text = parsedSourceCode.Substring(offset, styleInsertion.Index - offset);

                //var text = parsedSourceCode.Substring(styleInsertion.Index, styleInsertion.Scope.Length);

                this.AddTextElementStyle(styleInsertion, offset);

            }




        }

        private void AddTextElementStyle(TextInsertion TextInsertion, int offset)
        {
            var style = GetStyleForScope(TextInsertion.Scope);

            if (style != null)
            {
                // les caractère sont enlever '\n'
                int Start = offset + TextInsertion.Scope.Index + 1;
                int Length = TextInsertion.Scope.Length;

                TextElementStyle textElementStyle = new TextElementStyle(Start, Length);
                textElementStyle.IsBlod = style.Bold;
                textElementStyle.IsItalic = style.Italic;
                textElementStyle.FontColor = style.Foreground;


                this.TextStructure.TextElementStyles.Add(textElementStyle);

            }
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

    }
}
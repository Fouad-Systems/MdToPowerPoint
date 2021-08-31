using Microsoft.Toolkit.Parsers.Markdown;
using Microsoft.Toolkit.Parsers.Markdown.Blocks;
using Newtonsoft.Json;
using PowerPointLibrary.Entities;
using PowerPointLibrary.Exceptions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.BLO
{
    /// <summary>
    /// Create PresentationStructure instance from Markdown structure
    /// </summary>
    public class PresentationStructureBLO
    {
        #region attributes
        private PresentationStructure _PresentationStructure;

        //private PresentationStructureBLO _PresentationStructureBLO;
        private TemplateBLO _TemplateBLO;
        private CommentActionBLO _CommentActionBLO;
        private SlideBLO _SlideBLO;
        private SlideZoneBLO _SlideZoneBLO;
        private MarkdownBlockBLO _MarkdownBlockBLO;

        #endregion

        public PresentationStructureBLO(PresentationStructure presentationStructure)
        {
            _PresentationStructure = presentationStructure;

            // Init BLO
            _CommentActionBLO = new CommentActionBLO();
            _SlideBLO = new SlideBLO(_PresentationStructure);
            _SlideZoneBLO = new SlideZoneBLO();
            _TemplateBLO = new TemplateBLO(_PresentationStructure);
            _MarkdownBlockBLO = new MarkdownBlockBLO();

        }

        /// <summary>
        /// Crete the presentation data structure from mdDocument
        /// </summary>
        /// <param name="mdDocument"></param>
        public void CreatePresentationDataStructure(MarkdownDocument mdDocument)
        {
            // Amélioration
            // il faut d'abord, trouver le nombre des slides avec le nombre de type de contenue dans 
            // chaque slide
            // ensuite choisir la layout convenable pour chaque contenue 
            // ensuite read data frm mdDocument o PresentationDataStrucure

            int contentNumber = 0;
            foreach (var element in mdDocument.Blocks)
            {
                if (element.Type == MarkdownBlockType.YamlHeader) continue;
                // Create Slide if Header < 2
                // if header and header < 2
                if (element is HeaderBlock header)
                {
                    if (header.HeaderLevel <= 2)
                    {
                        // Set Default layout name
                        string layout = "";
                        if (header.HeaderLevel == 1) layout = "Titre session";
                        if (header.HeaderLevel >= 2) layout = "Titre contenu";

                        // Add new Slide
                        _SlideBLO.AddSlide(layout);
                        _SlideBLO.FindZoneForTitle();
                        contentNumber = 0;

                        // Add Text to TitleZone
                        SlideZoneStructure zoneTitle = this.CurrentSlide.CurrentZone;
                        if (zoneTitle != null)
                        {
                            if (zoneTitle.Text == null) zoneTitle.Text = new TextStructure();
                            _SlideZoneBLO.AddMarkdownBlockToSlideZone(zoneTitle, header);
                        }
                        continue;
                    }
                }

                // if paragraphe is action
                if (_MarkdownBlockBLO.isAction(element))
                {
                    string comment = _MarkdownBlockBLO.GetComment(element); ;
                    CommentAction commentAction =  _SlideBLO.ApplyAction(comment);
                    if (commentAction.ActionType == CommentAction.ActionTypes.NewSlide)
                        contentNumber = 0;
                    continue;
                }

                /// Insert Note 
                if (this.CurrentSlide.AddToNotes)
                {
                    _SlideBLO.AddNotes(element, ++contentNumber);
                    continue;
                }
                else
                {
                    if (! _MarkdownBlockBLO.isAction(element))
                        _SlideBLO.AddExplicationNotes(element, ++contentNumber);
                }

                // Indicate UseSlide
                if (this.CurrentSlide.UseSlideOrder != 0) continue;

              

                // Find Zone
                if (new MarkdownBlockBLO().IsImage(element))
                {
                    _SlideBLO.FindZoneForImage();

                }
                else
                {
                    _SlideBLO.FindZoneForTexte();
                }

                if (this.CurrentSlide.CurrentZone != null)
                {
                    if (this.CurrentSlide.CurrentZone.Text == null)
                        this.CurrentSlide.CurrentZone.Text = new TextStructure();

                    // return à la ligne si une nouvelle paragraphe est ajouté
                    int count_befor = this.CurrentSlide.CurrentZone.Text.Text.Count();
                    _SlideZoneBLO.AddMarkdownBlockToSlideZone(this.CurrentSlide.CurrentZone, element);
                    if (this.CurrentSlide.CurrentZone.Text.Text.Count() > count_befor)
                        this.CurrentSlide.CurrentZone.Text.Text += "\r";
                }
            }
        }


        public SlideStructure CurrentSlide
        {
            get
            {
                return this._PresentationStructure.CurrentSlide;
            }
        }

    }
}

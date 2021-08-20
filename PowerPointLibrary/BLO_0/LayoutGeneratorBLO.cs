using PowerPointLibrary.Entities;
using PowerPointLibrary.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.BLO
{
    /// <summary>
    /// Generate SlideZone 
    /// </summary>
    public class LayoutGeneratorBLO
    {

        public GLayoutStructure Parse(string layout)
        {
            GLayoutStructure gLayoutStructure = new GLayoutStructure();
            //  t 6 6
            //  t-2 6-3 6-9
            //  t-2 3-3 3-5 3-3
            //  t 7-5 5-5 12-3 p-30


            var parts = layout.Split(' ');

            foreach (string bloc in parts)
            {
                // Title
                if (bloc.StartsWith("t"))
                {
                    if (bloc.Count() > 1)
                    {
                        var title_parts = bloc.Split('-');

                        if (title_parts.Length != 2)
                        {
                            throw new PplException($"The title ligne '{bloc}' must be as t-2 ");
                        }
                        gLayoutStructure.TitleLines = Convert.ToInt32(title_parts[1]);

                    }
                    else
                    {
                        gLayoutStructure.TitleLines = 2;
                    }

                    continue;
                }

                // Padding
                if (bloc.StartsWith("p"))
                {
                    if (bloc.Count() > 1)
                    {
                        var padding_parts = bloc.Split('-');

                        if (padding_parts.Length != 2)
                        {
                            throw new PplException($"The padding  '{bloc}' must be as p-30 ");
                        }
                        gLayoutStructure.Padding = Convert.ToInt32(padding_parts[1]);
                    }
                    else
                    {
                        gLayoutStructure.Padding = 30;
                    }

                    continue;
                }



                var bloc_parts = bloc.Split('-');
                int cols = 0;
                int lines = 0;

                if (bloc_parts.Length == 1)
                {
                    cols = Convert.ToInt32(bloc_parts[0]);


                    lines = 12 - 1 - gLayoutStructure.TitleLines;
                }
                else
                {
                    cols = Convert.ToInt32(bloc_parts[0]);
                    lines = Convert.ToInt32(bloc_parts[1]);
                }

                if (cols > 12) throw new PplException($"The columns '{bloc}' mast be < 12");
                if (lines > 12) throw new PplException($"The lines '{bloc}' mast be < 12");

                this.AddBloc(gLayoutStructure, cols, lines);
            }


            return gLayoutStructure;
        }

        private void AddBloc(GLayoutStructure gLayoutStructure, int columns, int lines)
        {
            if (gLayoutStructure.Rows == null) gLayoutStructure.Rows = new List<GLayoutStructure.Row>();
            if (gLayoutStructure.Rows.Count == 0) gLayoutStructure.Rows.Add(new GLayoutStructure.Row());

            var lastRow = gLayoutStructure.Rows.Last();


            // Create or Use last Row 
            var free_columns = 12 - lastRow.Blocs.Sum(b => b.Columns);
            if (free_columns < columns) gLayoutStructure.Rows.Add(new GLayoutStructure.Row());

            lastRow = gLayoutStructure.Rows.Last();
            lastRow.Blocs.Add(new GLayoutStructure.Bloc(columns, lines));

            var SlideLines = gLayoutStructure.Rows.Sum(r => r.Blocs.Max(b => b.Lines));

            if ((SlideLines + gLayoutStructure.TitleLines) > 12)
            {
                throw new PplException($"The layout lines is more then 12");

            }


        }


        /// <summary>
        /// Generate Slide zone
        /// </summary>
        /// <param name="CurrentSlide"></param>
        /// <param name="gLayoutStructure"></param>
        public void GenerateSlideZone(SlideStructure CurrentSlide, GLayoutStructure gLayoutStructure)
        {
            SlideZoneBLO _SlideZoneStructureBLO = new SlideZoneBLO();
            float currentTop = 0;
            float currentLeft = 0;
            CurrentSlide.IsGenerated = true;

            // Default value 
            float ColumnHeight = 90;
            float ColumnWith = 160;
            float SlideHeight = 1080;

            if (CurrentSlide.GeneratedSlideZones == null)
                CurrentSlide.GeneratedSlideZones = new List<SlideZoneStructure>();

            // Add Title Zone
            if (gLayoutStructure.TitleLines > 0)
            {
                SlideZoneStructure TitleslideZoneStructure = new SlideZoneStructure();
                TitleslideZoneStructure.Order = 1;
                TitleslideZoneStructure.Name = "Title";
                TitleslideZoneStructure.ContentTypes.Add(Entities.Enums.ContentTypes.Title);
                TitleslideZoneStructure.Top = 0;
                TitleslideZoneStructure.Left = 0;
                TitleslideZoneStructure.Width = ColumnWith * 12;
                TitleslideZoneStructure.Height = ColumnHeight * 2;
                TitleslideZoneStructure.Row = 1;

                currentTop += TitleslideZoneStructure.Height;

                CurrentSlide.GeneratedSlideZones.Add(TitleslideZoneStructure);
            }

            int order = 1;
            int padding = gLayoutStructure.Padding;
            int RowNumber = 2;

            foreach (var row in gLayoutStructure.Rows)
            {

                foreach (var bloc in row.Blocs)
                {
                    SlideZoneStructure slideZoneStructure = new SlideZoneStructure();
                    slideZoneStructure.ContentTypes.Add(Entities.Enums.ContentTypes.Text);
                    slideZoneStructure.ContentTypes.Add(Entities.Enums.ContentTypes.Image);
                    slideZoneStructure.Order = ++order;
                    slideZoneStructure.ContentTypes.Add(Entities.Enums.ContentTypes.Text);
                    slideZoneStructure.ContentTypes.Add(Entities.Enums.ContentTypes.Image);

                    slideZoneStructure.GColumns = bloc.Columns;
                    slideZoneStructure.GLines = bloc.Lines;
                    slideZoneStructure.Width = bloc.Columns * ColumnWith - padding * 2;
                    slideZoneStructure.Height = bloc.Lines * ColumnHeight - padding * 2;
                    slideZoneStructure.Top = currentTop + padding;
                    slideZoneStructure.Left = currentLeft + padding;
                    slideZoneStructure.Row = RowNumber;
                    currentLeft += slideZoneStructure.Width + padding * 2;

                    CurrentSlide.GeneratedSlideZones.Add(slideZoneStructure);
                }

                RowNumber++;
                float row_Height = row.Blocs.Max(b => b.Lines) * ColumnHeight;
                currentTop += row_Height;
                currentLeft = 0;

            }

            // Center All zone, else Title
            int rowNumber = 2; // start after title
            foreach (var row in gLayoutStructure.Rows)
            {
                
                var zones = CurrentSlide.GeneratedSlideZones.Where(z => z.Row == rowNumber).ToList();
       
                float max_height = zones.Max(z => z.Height);

                foreach (var zone in zones)
                {
                    zone.Top += (max_height - zone.Height) / 3; // nombre d'or
                }
                rowNumber++;
            }

            // Copy TitleZone
            var TitleZone =  SlideBLO.GetTitleZoneFromSlideZones(CurrentSlide);

            if ( _SlideZoneStructureBLO.IsTitle(TitleZone))
            {
                CurrentSlide.CurrentZone = CurrentSlide.GeneratedSlideZones.First();
                CurrentSlide.CurrentZone.Text = TitleZone.Text.Clone() as TextStructure;
                CurrentSlide.CurrentZone.Name = TitleZone.Name;
            }

            // Center the content in the slide

            float ContentHeight = 0;
            foreach (var row in gLayoutStructure.Rows)
            {

                float row_Height = row.Blocs.Max(b => b.Lines) * ColumnHeight;
                ContentHeight += row_Height;
            }
            float freeHeigh = SlideHeight - ContentHeight - ColumnHeight * gLayoutStructure.TitleLines ;
            freeHeigh -= ColumnHeight;// footer line


            float top_offset = freeHeigh / 2;

            foreach (var generatedSlideZone in CurrentSlide.GeneratedSlideZones)
            {
                if (!_SlideZoneStructureBLO.IsTitle(generatedSlideZone))
                    generatedSlideZone.Top += top_offset;
            }


        }
    }
}

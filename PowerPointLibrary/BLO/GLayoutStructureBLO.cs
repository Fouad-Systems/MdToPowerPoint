using PowerPointLibrary.Entities;
using PowerPointLibrary.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.BLO
{
    public class GLayoutStructureBLO
    {

        public GLayoutStructure Parse(string layout)
        {
            GLayoutStructure gLayoutStructure = new GLayoutStructure();
            //  t 6 6
            //  t-2 6-3 6-9
            //  t-2 3-3 3-5 3-3


            var parts = layout.Split(' ');

            foreach (string bloc in parts)
            {
                if (bloc.StartsWith("t"))
                {
                    if(bloc.Count()> 1)
                    {
                        var title_parts = bloc.Split('-');
                        
                        if(title_parts.Length != 2)
                        {
                            throw new PplException($"The title ligne '{bloc}' must be as t-2 ");
                        }
                        gLayoutStructure.TitleLines = Convert.ToInt32(title_parts[1]);

                    }
                    else
                    {
                        gLayoutStructure.TitleLines = 1;
                    }
                }
                else
                {
                    
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

            if( (SlideLines + gLayoutStructure.TitleLines) > 12)
            {
                throw new PplException($"The layout lines is more then 12");

            }


        }


        public void GenerateSlideZone(SlideStructure CurrentSlide, GLayoutStructure gLayoutStructure)
        {

            int currentTop  = 0;
            int currentLeft = 0;
            CurrentSlide.IsGenerated = true;

            if (CurrentSlide.GeneratedSlideZones == null)
                CurrentSlide.GeneratedSlideZones = new List<SlideZoneStructure>();


            if (gLayoutStructure.TitleLines > 0)
            {
                SlideZoneStructure TitleslideZoneStructure = new SlideZoneStructure();
                TitleslideZoneStructure.Order = 1;
                TitleslideZoneStructure.Name = "Title";
                TitleslideZoneStructure.Top = 0;
                TitleslideZoneStructure.Left = 0;
                TitleslideZoneStructure.Width = 160 * 12;
                TitleslideZoneStructure.Height = 90 * 2;
                TitleslideZoneStructure.Row = 1;

                currentTop += TitleslideZoneStructure.Height;

                CurrentSlide.GeneratedSlideZones.Add(TitleslideZoneStructure);
            }

            int order = 1;
            int padding = 30;
            int RowNumber = 2;

            foreach (var row in gLayoutStructure.Rows)
            {
               
                foreach (var bloc in row.Blocs)
                {
                    SlideZoneStructure slideZoneStructure = new SlideZoneStructure();

                    slideZoneStructure.Order = ++order;
                    slideZoneStructure.ContentTypes.Add(Entities.Enums.ContentTypes.Text);
                    slideZoneStructure.ContentTypes.Add(Entities.Enums.ContentTypes.Image);

                    slideZoneStructure.GColumns = bloc.Columns;
                    slideZoneStructure.GLines = bloc.Lines;
                    slideZoneStructure.Width = bloc.Columns * 160 - padding * 2;
                    slideZoneStructure.Height = bloc.Lines * 90 - padding * 2;
                    slideZoneStructure.Top = currentTop + padding;
                    slideZoneStructure.Left = currentLeft + padding;
                    slideZoneStructure.Row = RowNumber;
                    currentLeft += slideZoneStructure.Width + padding * 2;

                    CurrentSlide.GeneratedSlideZones.Add(slideZoneStructure);
                }

                RowNumber++;
                int row_Height = row.Blocs.Max(b => b.Lines) * 90;
                currentTop += row_Height;
                currentLeft = 0;


             

               




            }


            int rowNumber = 1;
            foreach (var row in gLayoutStructure.Rows)
            {
                var zones = CurrentSlide.GeneratedSlideZones.Where(z => z.Row == rowNumber).ToList();
                rowNumber++;
                if (zones.Count() == 1) continue;

                int max_height = zones.Max(z => z.Height);

                foreach (var zone in zones)
                {
                    zone.Top += (max_height - zone.Height) / 2;
                }


               
            }

             
          




            // this.CopySlideZoneToGeneratedSlideZone();

            // Copy TitleZone
            var TitleZone = CurrentSlide.CurrentZone;

            if (CurrentSlide.CurrentZone.Name == "Title" || CurrentSlide.CurrentZone.Name == "Titre")
            {
                CurrentSlide.CurrentZone = CurrentSlide.GeneratedSlideZones.First();
                CurrentSlide.CurrentZone.Text = TitleZone.Text.Clone() as TextStructure;
                CurrentSlide.CurrentZone.Name = TitleZone.Name;

            }



        }
    }
}

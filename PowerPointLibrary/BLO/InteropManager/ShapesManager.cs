﻿namespace PowerPointLibrary.Manager
{
    using System.Collections.Generic;
    using System.Linq;

    using PowerPointLibrary.Helper.Contracts;
    using PowerPointLibrary.Helper.Objects;


    using PPT = Microsoft.Office.Interop.PowerPoint;
    using OFFICE = Microsoft.Office.Core;
    using System.Drawing;

    public class ShapesManager : IShapesManager
    {

        public ShapesManager()
        {
            
        }

        /// <summary>
        ///     Add an existing picture to a slide
        /// </summary>
        /// <param name="slide">PPT.Slide object instance</param>
        /// <param name="file">The file path and name of the image</param>
        /// <param name="leftPosition">x Location</param>
        /// <param name="topPosition">y Location</param>
        /// <param name="width">Width of the picture</param>
        /// <param name="height">Height of the picture</param>
        /// <returns></returns>
        public PPT.Shape AddPicture(
                PPT.Slide slide,
                string file,
                float leftPosition,
                float topPosition,
                float width,
                float height)
        {

            // add a shape then add picture 


          //  PPT.Shape shapeOut1 = slide.Shapes.AddShape(OFFICE.MsoAutoShapeType.msoShapeRectangle, leftPosition, topPosition, width, height);



            PPT.Shape shapeOut = slide.Shapes.AddPicture(
                    file,
                    OFFICE.MsoTriState.msoTrue,
                    OFFICE.MsoTriState.msoTrue,
                    leftPosition,
                    topPosition,
                    width,
                    height);

           // shapeOut.Line.Weight = 1;
           // shapeOut.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(ColorTranslator.FromHtml("#aaa"));
           //// shapeOut.Shadow.Type = OFFICE.MsoShadowType.msoShadow21;

          


            return shapeOut;
        }

        /// <summary>
        ///     Add a table to a slide
        /// </summary>
        /// <param name="slide">PPT.Slide object instance</param>
        /// <param name="numRows">Number of rows to create in the table</param>
        /// <param name="numColumns">Number of columns to create in the table</param>
        /// <param name="xLocation">x location</param>
        /// <param name="yLocation">y location</param>
        /// <param name="width">Table shape width</param>
        /// <param name="height">Table shape height</param>
        /// <returns></returns>
        public PPT.Shape AddTableToSlide(
                PPT.Slide slide,
                int numRows,
                int numColumns,
                float xLocation,
                float yLocation,
                float width,
                float height)
        {
            PPT.Shape table = slide.Shapes.AddTable(numRows, numColumns, xLocation, yLocation, width, height);
            return table;
        }

        /// <summary>
        ///     Add a Textbox to a slide
        /// </summary>
        /// <param name="slide">PPT.Slide object to add a textbox to</param>
        /// <param name="orientation">Orientation of the textbox</param>
        /// <param name="xLocation">x Location</param>
        /// <param name="yLocation">y Location</param>
        /// <param name="width">Textbox width</param>
        /// <param name="height">Textbox height</param>
        /// <returns></returns>
        public PPT.Shape AddTextBoxToSlide(
                PPT.Slide slide,
                OFFICE.MsoTextOrientation orientation,
                float xLocation,
                float yLocation,
                float width,
                float height)
        {
            PPT.Shape textbox = slide.Shapes.AddTextbox(orientation, xLocation, yLocation, width, height);
            return textbox;
        }

        /// <summary>
        ///     Draws a line on a slide
        /// </summary>
        /// <param name="slide">PPT.Slide object instance</param>
        /// <param name="xStartLocation">Starting x location</param>
        /// <param name="xEndLocation">Ending x location</param>
        /// <param name="yStartLocation">Starting y location</param>
        /// <param name="yEndLocation">Ending y location</param>
        /// <returns></returns>
        public PPT.Shape DrawLine(
                PPT.Slide slide,
                float xStartLocation,
                float xEndLocation,
                float yStartLocation,
                float yEndLocation)
        {
            return slide.Shapes.AddLine(xStartLocation, yStartLocation, xEndLocation, yEndLocation);
        }

        /// <summary>
        ///     Draw a shape on a slide
        /// </summary>
        /// <param name="slide">PPT.Slide object instance</param>
        /// <param name="shapeType">The shape type. Any shape from the MsoAutoShapeType can be specified</param>
        /// <param name="leftPosition">x position</param>
        /// <param name="topPosition">y position</param>
        /// <param name="width">Shape width</param>
        /// <param name="height">Shape height</param>
        /// <returns></returns>
        public PPT.Shape DrawShape(
                PPT.Slide slide,
                OFFICE.MsoAutoShapeType shapeType,
                float leftPosition,
                float topPosition,
                float width,
                float height)
        {
            return slide.Shapes.AddShape(shapeType, leftPosition, topPosition, width, height);
        }

        /// <summary>
        ///     Find all shapes of a type in a presentation
        /// </summary>
        /// <param name="presentation">PPT.Presentation object instance</param>
        /// <param name="shapeType">The shape type to look for</param>
        /// <returns></returns>
        public List<ShapesofType> FindShapesInPresentation(PPT.Presentation presentation, OFFICE.MsoAutoShapeType shapeType)
        {
            return (from PPT.Slide slide in presentation.Slides
                    from PPT.Shape shape in slide.Shapes
                    where shape.AutoShapeType == shapeType
                    select new ShapesofType { shape = shape, shapeType = shape.Type, slide = slide }).ToList();
        }

        /// <summary>
        ///     Set the text in the textbox
        /// </summary>
        /// <param name="textbox">PPT.Shape that is a textbox</param>
        /// <param name="text">Text</param>
        public void SetTextBoxText(PPT.Shape textbox, string text)
        {
            textbox.TextEffect.Text = text;
        }

        /// <summary>
        /// Add a web hyperlink to any shape
        /// </summary>
        /// <param name="shape">shape in</param>
        /// <param name="hyperLinkUrl">string url such as "http://google.com"</param>
        public void AddHyperLinkToWebsite(PPT.Shape shape, string hyperLinkUrl)
        {
            shape.ActionSettings[PPT.PpMouseActivation.ppMouseClick].Hyperlink.Address = hyperLinkUrl;
        }

        /// <summary>
        /// Add an action to carry out when a shape is clicked with the mouse
        /// </summary>
        /// <param name="shape">shape in</param>
        /// <param name="action">the action to carry out when the shape is clicked</param>
        public void AddClickedActionToShape(PPT.Shape shape, PPT.PpActionType action)
        {
            shape.ActionSettings[PPT.PpMouseActivation.ppMouseClick].Action = action;
        }

        /// <summary>
        /// Add an action to carry out when a shape is clicked with the mouse
        /// </summary>
        /// <param name="shape">shape in</param>
        /// <param name="action">the action to carry out when the shape is clicked</param>
        public void AddMouseOverActionToShape(PPT.Shape shape, PPT.PpActionType action)
        {
            shape.ActionSettings[PPT.PpMouseActivation.ppMouseOver].Action = action;
        }
    }
}
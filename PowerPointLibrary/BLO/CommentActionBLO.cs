using PowerPointLibrary.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.BLO
{
    public class CommentActionBLO
    {
        /// <summary>
        /// Test the comment is a valide action
        /// </summary>
        /// <param name="comment">the comment to test</param>
        /// <returns></returns>
        public bool IsAction(string comment)
        {
            // The possible actions
            // <!-- layout : Titre et contenu --> : change layout
            // <!-- zone : Colone 1 --> : define zone
            // <!-- note --> : write to zone to notes
            // <!-- end note --> : switch to wrtie to zone
            // <!-- new slide : Titre et contenu -->
            // <!-- new slide -->
            // <!-- use : slide 1 --> : use Slide 1 that exist in the file : OutputfileName.slides.pptx
            // <!-- g layout : t 6-3 6-9 --> : Geneate layout, t: titre, 6-3 ligne 1, 6-9 ligne 2

            if (
                comment.StartsWith("<!-- layout :", true, null) ||
                comment.StartsWith("<!-- zone :", true, null) ||
                comment.StartsWith("<!-- note -->", true, null) ||
                comment.StartsWith("<!-- end note -->", true, null) ||
                comment.StartsWith("<!-- new slide", true, null) ||
                comment.StartsWith("<!-- g layout :", true, null)
                )
                return true;
            else
                return false;

        }

        public CommentAction ParseComment(string comment)
        {
           
            if (!this.IsAction(comment)) return null;

            CommentAction commentAction = null;

            if (comment.StartsWith("<!-- layout :", true, null))
            {
                commentAction = new CommentAction(comment);
                commentAction.ActionType = CommentAction.ActionTypes.ChangeLayout;
                string layout = comment.Replace("<!-- layout :", "").Replace("-->", "").Trim();
                commentAction.Layout = layout;
            }
            if (comment.StartsWith("<!-- zone :", true, null))
            {
                commentAction = new CommentAction(comment);
                commentAction.ActionType = CommentAction.ActionTypes.ChangeZone;
                string zoneName = comment.Replace("<!-- zone :", "").Replace("-->", "").Trim();
                commentAction.ZoneName = zoneName;
            }
            if (comment.StartsWith("<!-- new slide", true, null))
            {
                commentAction = new CommentAction(comment);
                commentAction.ActionType = CommentAction.ActionTypes.NewSlide;
                string Layout = comment.Replace("<!-- new slide :", "").Replace("-->", "").Trim();
                commentAction.Layout = Layout;
            }
            if (comment.StartsWith("<!-- note -->", true, null))
            {
                commentAction = new CommentAction(comment);
                commentAction.ActionType = CommentAction.ActionTypes.Note;
                commentAction.ZoneName = null;
            }
            if (comment.StartsWith("<!-- end note", true, null))
            {
                commentAction = new CommentAction(comment);
                commentAction.ActionType = CommentAction.ActionTypes.EndNote;
                commentAction.ZoneName = null;
            }
            if (comment.StartsWith("<!-- use : slide", true, null))
            {
                commentAction = new CommentAction(comment);
                commentAction.ActionType = CommentAction.ActionTypes.UseSlide;
                string stringSlideOrder = comment.Replace("<!-- use : slide", "").Replace("-->", "");

                int SlideOrder = Convert.ToInt32(stringSlideOrder.Trim());

                commentAction.UseSlideOrder = SlideOrder;
            }
            if (comment.StartsWith("<!-- g layout :", true, null))
            {
                commentAction = new CommentAction(comment);
                commentAction.ActionType = CommentAction.ActionTypes.GenerateLayout;
                string layout = comment.Replace("<!-- g layout :", "").Replace("-->", "").Trim();
                commentAction.GLayoutStructure = new GLayoutStructureBLO().Parse(layout);
            }


            return commentAction;
        }
    }
}

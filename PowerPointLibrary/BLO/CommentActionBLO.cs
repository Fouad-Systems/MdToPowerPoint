﻿using PowerPointLibrary.Entities;
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
            // <!-- zone : Notes --> : change zone to notes
            // <!-- new slide : Titre et contenu -->
            // <!-- new slide -->

            if (
                comment.StartsWith("<!-- layout :", true, null) ||
                comment.StartsWith("<!-- zone :", true, null) ||
                comment.StartsWith("<!-- new slide", true, null)
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



            return commentAction;
        }
    }
}
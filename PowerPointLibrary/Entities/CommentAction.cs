using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.Entities
{
    /// <summary>
    /// Comment action 

    /// </summary>
    public class CommentAction
    {

        public enum ActionTypes
        {
            Empty,
            ChangeLayout,
            ChangeZone,
            NewSlide
        }

        public string Comment { get; set; }


        public ActionTypes ActionType { get; set; }
        public string Layout { get; internal set; }
        public string ZoneName { get; internal set; }

        public CommentAction(string Comment)
        {
            this.Comment = Comment;


        }



    }
}

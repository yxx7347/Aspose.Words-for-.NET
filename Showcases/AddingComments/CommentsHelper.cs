using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;

namespace AddingComments
{
    public class CommentsHelper
    {
        public static Comment AddComment(Document doc, string authorName, string initials, DateTime dateTime,
            string commentText)
        {
            DocumentBuilder builder = new DocumentBuilder(doc);

            Comment comment = new Comment(doc, authorName, initials, dateTime);
            comment.SetText(commentText);

            builder.CurrentParagraph.AppendChild(comment);

            return comment;
        }

        public static List<Comment> ExtractComments(Document doc)
        {
            List<Comment> collectedComments = new List<Comment>();

            // Collect all comments in the document
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            // Look through all comments and gather information about those written by the authorName author.
            foreach (Comment comment in comments)
            {
                collectedComments.Add(comment);
            }

            return collectedComments;
        }
        
        public static List<Comment> ExtractComments(Document doc, string authorName)
        {
            List<Comment> collectedComments = new List<Comment>();

            // Collect all comments in the document
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            // Look through all comments and gather information about those written by the authorName author.
            foreach (Comment comment in comments)
            {
                if (comment.Author == authorName)
                    collectedComments.Add(comment);
            }

            return collectedComments;
        }

        public static void RemoveComments(Document doc)
        {
            // Collect all comments in the document
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            // Remove all comments.
            comments.Clear();
        }

        public static void RemoveComments(Document doc, string authorName)
        {
            // Collect all comments in the document
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            // Look through all comments and remove those written by the authorName author.
            for (int i = comments.Count - 1; i >= 0; i--)
            {
                Comment comment = (Comment)comments[i];
                if (comment.Author == authorName)
                    comment.Remove();
            }
        }

        public static void MarkCommentsAsDone(Document doc)
        {
            // Collect all comments in the document
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            foreach (Comment comment in comments)
            {
                if (comment.Ancestor == null)
                {
                    comment.Done = true;
                }
            }
        }

        public static void MarkCommentAsDone(Comment comment)
        {
            comment.Done = true;
        }
    }
}

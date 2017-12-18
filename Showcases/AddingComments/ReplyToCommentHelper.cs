// INSP DD: Aspose Copyright etc.
using System;
using System.Collections.Generic;
using Aspose.Words;

namespace AddingComments
{

    public static class ReplyToCommentHelper
    {
        // INSP DD: this may be too trivial to create a method, remove
        public static void AddReplyToComment(Comment comment, string authorName, string initials, DateTime dateTime, string commentText)
        {
            comment.AddReply(authorName, initials, dateTime, commentText);
        }

        public static List<Comment> ExtractReplies(Document doc)
        {
            List<Comment> collectedComments = new List<Comment>();

            // Collect all comments in the document
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            foreach (Comment comment in comments)
            {
                foreach (Comment reply in comment.Replies)
                {
                    collectedComments.Add(reply);
                }
            }

            return collectedComments;
        }

        public static List<Comment> ExtractReplies(Comment comment)
        {
            List<Comment> collectedComments = new List<Comment>();

            foreach (Comment reply in comment.Replies)
            {
                collectedComments.Add(reply);
            }

            return collectedComments;
        }

        // INSP DD: Usually when index is passed the methods name contains At e.g. RemoveReplyAt
        public static void RemoveReply(Comment comment, int replyIndex)
        {
            comment.RemoveReply(comment.Replies[replyIndex]);
        }

        public static void RemoveReplies(Document doc)
        {
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            foreach (Comment comment in comments)
            {
                if (comment.Ancestor != null)
                {
                    comment.Remove();
                }
            }
        }

        // INSP DD: this may be too trivial to create a method, remove
        public static void RemoveReplies(Comment comment)
        {
            comment.RemoveAllReplies();
        }

        public static bool IsReply(Comment comment)
        {
            if (comment.Ancestor != null)
            {
                return true;
            }

            return false;
        }

        // INSP DD: unused, either use in some usecase or remove 
        public static void MarkRepliesAsDone(Comment comment)
        {
            foreach (Comment reply in comment.Replies)
            {
                reply.Done = true;
            }
        }

        // INSP DD: this may be too trivial to create a method, remove
        public static void MarkReplyAsDone(Comment reply)
        {
            reply.Done = true;
        }
    }
}

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
using NUnit.Framework;

namespace AddingComments
{
    [TestFixture]
    public class Program
    {
        // Sample infrastructure.
        static readonly string ExeDir = Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath) + Path.DirectorySeparatorChar;
        static readonly string DataDir = new Uri(new Uri(ExeDir), @"../../Data/").LocalPath;

        [Test]
        public static void Main()
        {
            Comment();
            ReplyToComment();
        }

        public static void Comment()
        {
            // Create new test document for adding comments and replies.
            Document doc = new Document();

            // Add test comments to the document
            for (int i = 0; i <= 10; i++)
            {
                CommentsHelper.AddComment(doc, "Author " + i, "Initials " + i, DateTime.Now, "Comment text " + i);
            }

            Console.WriteLine("Comments are added!");

            // Extract the information about the comments of all the authors.
            foreach (Comment comment in CommentsHelper.ExtractComments(doc))
                Console.Write(comment.GetText());

            // Remove comments by the "Author 2" author.
            CommentsHelper.RemoveComments(doc, "Author 2");
            Console.WriteLine("Comments are removed!");

            // Extract the information about the comments of the "Author 1" author and mark as "Done"
            foreach (Comment comment in CommentsHelper.ExtractComments(doc, "Author 1"))
                CommentsHelper.MarkCommentAsDone(comment);

            // Mark all comments in the document as "Done"
            CommentsHelper.MarkCommentsAsDone(doc);

            Console.WriteLine("All comments marks as 'Done'");

            // Remove all comments.
            CommentsHelper.RemoveComments(doc);
            Console.WriteLine("All comments are removed!");

            // Save the document.
            doc.Save(DataDir + "Test File Comment Out.docx");
        }

        public static void ReplyToComment()
        {
            // Create new test document for adding comments and replies.
            Document doc = new Document();

            // Add comments with replies to the document
            for (int i = 0; i <= 10; i++)
            {
                Comment comment = CommentsHelper.AddComment(doc, "Author " + i, "Initials " + i, DateTime.Now, "Comment text " + i);

                for (int y = 0; y <= 10; y++)
                {
                    ReplyToCommentHelper.AddReplyToComment(comment, "Reply author " + y, "Reply initials " + y,
                        DateTime.Now, "Reply to comment " + y);
                }
            }

            Console.WriteLine("All comments and replies are added!");

            // Extract the information about the all replies of all comments in the document
            foreach (Comment comment in ReplyToCommentHelper.ExtractReplies(doc))
            {
                Console.Write(comment.Ancestor.GetText());
                Console.Write(comment.GetText());
            }

            // Remove reply to comment by index.
            foreach (Comment comment in CommentsHelper.ExtractComments(doc, "Author 1"))
                ReplyToCommentHelper.RemoveReply(comment, 2);

            Console.WriteLine("Reply was removed!");

            // Extract the information about the replies from comment of the "Author 2" author and mark as "Done"
            foreach (Comment comment in CommentsHelper.ExtractComments(doc, "Author 2"))
                foreach (Comment reply in ReplyToCommentHelper.ExtractReplies(comment))
                ReplyToCommentHelper.MarkReplyAsDone(reply);

            // Mark replies of the "Author 1" comment author as "Done"
            foreach (Comment comment in CommentsHelper.ExtractComments(doc, "Author 3"))
                CommentsHelper.MarkCommentAsDone(comment);

            Console.WriteLine("All replies marks as 'Done'");

            // Remove reply of the "Author 4" comment author by index
            foreach (Comment comment in CommentsHelper.ExtractComments(doc, "Author 4"))
                ReplyToCommentHelper.RemoveReply(comment, 1);

            // Remove all replies of the "Author 4" comment author
            foreach (Comment comment in CommentsHelper.ExtractComments(doc, "Author 4"))
                ReplyToCommentHelper.RemoveReplies(comment);

            Console.WriteLine("Specific replies are removed!");

            // Check that comment is reply to
            foreach (Comment comment in CommentsHelper.ExtractComments(doc, "Author 5"))
                ReplyToCommentHelper.IsReply(comment);

            // Remove all replies to comments from the document
            ReplyToCommentHelper.RemoveReplies(doc);

            Console.WriteLine("All replies are removed!");

            // Save the document.
            doc.Save(DataDir + "Test File ReplyToComment Out.docx");
        }
    }
}

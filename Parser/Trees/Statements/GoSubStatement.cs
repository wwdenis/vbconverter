using System;
using System.Collections.Generic;
using System.Text;

using VBConverter.CodeParser.Trees.Comments;
using VBConverter.CodeParser.Trees.Names;

namespace VBConverter.CodeParser.Trees.Statements
{
    /// <summary>
    /// A parse tree for a GoSub statement (VB6).
    /// </summary>
    public sealed class GoSubStatement : LabelReferenceStatement
    {
        public string ReturnLabel
        {
            get
            {
                string label = string.Empty;

                if (Name != null && !string.IsNullOrEmpty(Name.Name))
                    label = string.Format("Return_{0}", Name.Name);

                return label;
            }
        }
        
        /// <summary>
        /// Constructs a parse tree for a GoSub statement.
        /// </summary>
        /// <param name="name">The label to branch to, if any.</param>
        /// <param name="isLineNumber">Whether the label is a line number.</param>
        /// <param name="span">The location of the parse tree.</param>
        /// <param name="comments">The comments for the parse tree.</param>
        public GoSubStatement(SimpleName name, bool isLineNumber, Span span, IList<Comment> comments) : base(TreeType.GoSubStatement, name, isLineNumber, span, comments)
        {
        }
    }
}

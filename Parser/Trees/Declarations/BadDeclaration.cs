using System;
using System.Collections.Generic;
using System.Text;

using VBConverter.CodeParser.Trees.Comments;

namespace VBConverter.CodeParser.Trees.Declarations
{
	/// <summary>
	/// A parse tree for an bad declaration.
	/// </summary>
	public sealed class BadDeclaration : Declaration
	{

		/// <summary>
		/// Constructs a new parse tree for an empty declaration.
		/// </summary>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
        public BadDeclaration(Span span, IList<Comment> comments): base(TreeType.SyntaxError, span, comments)
		{
		}
	}
}

using System;
using System.Collections.Generic;
using System.Text;
//
// Visual Basic .NET Parser
//
// Copyright (C) 2005, Microsoft Corporation. All rights reserved.
//
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER 
// EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF 
// MERCHANTIBILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
//
using VBConverter.CodeParser.Trees.Comments;
using VBConverter.CodeParser.Trees.Expressions;

namespace VBConverter.CodeParser.Trees.Statements
{
	/// <summary>
	/// A parse tree for a While block statement.
	/// </summary>
	public sealed class WhileBlockStatement : ExpressionBlockStatement
	{

		/// <summary>
		/// Constructs a new parse tree for a While statement block.
		/// </summary>
		/// <param name="expression">The expression.</param>
		/// <param name="statements">The statements in the block.</param>
		/// <param name="endStatement">The End statement for the block, if any.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public WhileBlockStatement(Expression expression, StatementCollection statements, EndBlockStatement endStatement, Span span, IList<Comment> comments) : base(TreeType.WhileBlockStatement, expression, statements, endStatement, span, comments)
		{
		}
	}
}

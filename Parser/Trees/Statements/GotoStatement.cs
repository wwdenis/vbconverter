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
using VBConverter.CodeParser.Trees.Names;

namespace VBConverter.CodeParser.Trees.Statements
{
	/// <summary>
	/// A parse tree for a GoTo statement.
	/// </summary>
	public sealed class GotoStatement : LabelReferenceStatement
	{

		/// <summary>
		/// Constructs a parse tree for a GoTo statement.
		/// </summary>
		/// <param name="name">The label to branch to, if any.</param>
		/// <param name="isLineNumber">Whether the label is a line number.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public GotoStatement(SimpleName name, bool isLineNumber, Span span, IList<Comment> comments) : base(TreeType.GotoStatement, name, isLineNumber, span, comments)
		{
		}
	}
}

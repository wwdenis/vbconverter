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

namespace VBConverter.CodeParser.Trees.Statements
{
	/// <summary>
	/// A parse tree for an Else If block statement.
	/// </summary>
	public sealed class ElseIfBlockStatement : BlockStatement
	{

		private readonly ElseIfStatement _ElseIfStatement;

		/// <summary>
		/// The Else If statement.
		/// </summary>
		public ElseIfStatement ElseIfStatement {
			get { return _ElseIfStatement; }
		}

		/// <summary>
		/// Constructs a new parse tree for an Else If block statement.
		/// </summary>
		/// <param name="elseIfStatement">The Else If statement.</param>
		/// <param name="statements">The statements in the block.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public ElseIfBlockStatement(ElseIfStatement elseIfStatement, StatementCollection statements, Span span, IList<Comment> comments) : base(TreeType.ElseIfBlockStatement, statements, span, comments)
		{

			if (elseIfStatement == null)
			{
				throw new ArgumentNullException("elseIfStatement");
			}

			SetParent(elseIfStatement);
			_ElseIfStatement = elseIfStatement;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, ElseIfStatement);
			base.GetChildTrees(childList);
		}
	}
}

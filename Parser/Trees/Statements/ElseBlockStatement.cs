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
	/// A parse tree for an Else block statement.
	/// </summary>
	public sealed class ElseBlockStatement : BlockStatement
	{

		private readonly ElseStatement _ElseStatement;

		/// <summary>
		/// The Else or Else If statement.
		/// </summary>
		public ElseStatement ElseStatement {
			get { return _ElseStatement; }
		}

		/// <summary>
		/// Constructs a new parse tree for an Else block statement.
		/// </summary>
		/// <param name="elseStatement">The Else statement.</param>
		/// <param name="statements">The statements in the block.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public ElseBlockStatement(ElseStatement elseStatement, StatementCollection statements, Span span, IList<Comment> comments) : base(TreeType.ElseBlockStatement, statements, span, comments)
		{

			if (elseStatement == null)
			{
				throw new ArgumentNullException("elseStatement");
			}

			SetParent(elseStatement);
			_ElseStatement = elseStatement;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, ElseStatement);
			base.GetChildTrees(childList);
		}
	}
}

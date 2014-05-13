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
	/// A parse tree for a Finally block statement.
	/// </summary>
	public sealed class FinallyBlockStatement : BlockStatement
	{

		private readonly FinallyStatement _FinallyStatement;

		/// <summary>
		/// The Finally statement.
		/// </summary>
		public FinallyStatement FinallyStatement {
			get { return _FinallyStatement; }
		}

		/// <summary>
		/// Constructs a new parse tree for a Finally block statement.
		/// </summary>
		/// <param name="finallyStatement">The Finally statement.</param>
		/// <param name="statements">The statements in the block.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public FinallyBlockStatement(FinallyStatement finallyStatement, StatementCollection statements, Span span, IList<Comment> comments) : base(TreeType.FinallyBlockStatement, statements, span, comments)
		{

			if (finallyStatement == null)
			{
				throw new ArgumentNullException("finallyStatement");
			}

			SetParent(finallyStatement);
			_FinallyStatement = finallyStatement;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, FinallyStatement);
			base.GetChildTrees(childList);
		}
	}
}

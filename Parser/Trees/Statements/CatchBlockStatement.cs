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
	/// A parse tree for a Catch block statement.
	/// </summary>
	public sealed class CatchBlockStatement : BlockStatement
	{

		private readonly CatchStatement _CatchStatement;

		/// <summary>
		/// The Catch statement.
		/// </summary>
		public CatchStatement CatchStatement {
			get { return _CatchStatement; }
		}

		/// <summary>
		/// Constructs a new parse tree for a Catch block statement.
		/// </summary>
		/// <param name="catchStatement">The Catch statement.</param>
		/// <param name="statements">The statements in the block.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public CatchBlockStatement(CatchStatement catchStatement, StatementCollection statements, Span span, IList<Comment> comments) : base(TreeType.CatchBlockStatement, statements, span, comments)
		{

			if (catchStatement == null)
			{
				throw new ArgumentNullException("catchStatement");
			}

			SetParent(catchStatement);
			_CatchStatement = catchStatement;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, CatchStatement);
			base.GetChildTrees(childList);
		}
	}
}

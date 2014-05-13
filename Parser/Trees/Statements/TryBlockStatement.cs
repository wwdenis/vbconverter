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
	/// A parse tree for a Try statement.
	/// </summary>
	public sealed class TryBlockStatement : BlockStatement
	{

		private readonly StatementCollection _CatchBlockStatements;
		private readonly FinallyBlockStatement _FinallyBlockStatement;
		private readonly EndBlockStatement _EndStatement;

		/// <summary>
		/// The Catch statements.
		/// </summary>
		public StatementCollection CatchBlockStatements {
			get { return _CatchBlockStatements; }
		}

		/// <summary>
		/// The Finally statement, if any.
		/// </summary>
		public FinallyBlockStatement FinallyBlockStatement {
			get { return _FinallyBlockStatement; }
		}

		/// <summary>
		/// The End Try statement, if any.
		/// </summary>
		public EndBlockStatement EndStatement {
			get { return _EndStatement; }
		}

		/// <summary>
		/// Constructs a new parse tree for a Try statement.
		/// </summary>
		/// <param name="statements">The statements in the Try block.</param>
		/// <param name="catchBlockStatements">The Catch statements.</param>
		/// <param name="finallyBlockStatement">The Finally statement, if any.</param>
		/// <param name="endStatement">The End Try statement, if any.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments of the parse tree.</param>
		public TryBlockStatement(StatementCollection statements, StatementCollection catchBlockStatements, FinallyBlockStatement finallyBlockStatement, EndBlockStatement endStatement, Span span, IList<Comment> comments) : base(TreeType.TryBlockStatement, statements, span, comments)
		{

			SetParent(catchBlockStatements);
			SetParent(finallyBlockStatement);
			SetParent(endStatement);

			_CatchBlockStatements = catchBlockStatements;
			_FinallyBlockStatement = finallyBlockStatement;
			_EndStatement = endStatement;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			base.GetChildTrees(childList);
			AddChild(childList, CatchBlockStatements);
			AddChild(childList, FinallyBlockStatement);
			AddChild(childList, EndStatement);
		}
	}
}

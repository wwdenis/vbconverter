using System;
using System.Collections.Generic;
using System.Diagnostics;
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
	/// A parse tree for an expression block statement.
	/// </summary>
	public abstract class ExpressionBlockStatement : BlockStatement
	{

		private readonly Expression _Expression;
		private readonly EndBlockStatement _EndStatement;

		/// <summary>
		/// The expression.
		/// </summary>
		public Expression Expression {
			get { return _Expression; }
		}

		/// <summary>
		/// The End statement for the block, if any.
		/// </summary>
		public EndBlockStatement EndStatement {
			get { return _EndStatement; }
		}

		protected ExpressionBlockStatement(TreeType type, Expression expression, StatementCollection statements, EndBlockStatement endStatement, Span span, IList<Comment> comments) : base(type, statements, span, comments)
		{

			Debug.Assert(type == TreeType.WithBlockStatement || type == TreeType.SyncLockBlockStatement || type == TreeType.WhileBlockStatement || type == TreeType.UsingBlockStatement);

			SetParent(expression);
			SetParent(endStatement);

			_Expression = expression;
			_EndStatement = endStatement;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, Expression);
			base.GetChildTrees(childList);
			AddChild(childList, EndStatement);
		}
	}
}

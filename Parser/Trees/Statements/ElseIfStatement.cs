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
	/// A parse tree for an Else If statement.
	/// </summary>
	public sealed class ElseIfStatement : Statement
	{

		private readonly Expression _Expression;
		private readonly Location _ThenLocation;

		/// <summary>
		/// The conditional expression.
		/// </summary>
		public Expression Expression {
			get { return _Expression; }
		}

		/// <summary>
		/// The location of the 'Then', if any.
		/// </summary>
		public Location ThenLocation {
			get { return _ThenLocation; }
		}

		/// <summary>
		/// Constructs a new parse tree for an Else If statement.
		/// </summary>
		/// <param name="expression">The conditional expression.</param>
		/// <param name="thenLocation">The location of the 'Then', if any.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public ElseIfStatement(Expression expression, Location thenLocation, Span span, IList<Comment> comments) : base(TreeType.ElseIfStatement, span, comments)
		{

			if (expression == null)
			{
				throw new ArgumentNullException("expression");
			}

			SetParent(expression);
			_Expression = expression;
			_ThenLocation = thenLocation;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, Expression);
		}
	}
}

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
	/// A parse tree for a Loop statement.
	/// </summary>
	public sealed class LoopStatement : Statement
	{

		private readonly bool _IsWhile;
		private readonly Location _WhileOrUntilLocation;
		private readonly Expression _Expression;

		/// <summary>
		/// Whether the Loop has a While or Until.
		/// </summary>
		public bool IsWhile {
			get { return _IsWhile; }
		}

		/// <summary>
		/// The location of the While or Until, if any.
		/// </summary>
		public Location WhileOrUntilLocation {
			get { return _WhileOrUntilLocation; }
		}

		/// <summary>
		/// The loop expression, if any.
		/// </summary>
		public Expression Expression {
			get { return _Expression; }
		}

		/// <summary>
		/// Constructs a parse tree for a Loop statement.
		/// </summary>
		/// <param name="expression">The loop expression, if any.</param>
		/// <param name="isWhile">WHether the Loop has a While or Until.</param>
		/// <param name="whileOrUntilLocation">The location of the While or Until, if any.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public LoopStatement(Expression expression, bool isWhile, Location whileOrUntilLocation, Span span, IList<Comment> comments) : base(TreeType.LoopStatement, span, comments)
		{

			SetParent(expression);

			_Expression = expression;
			_IsWhile = isWhile;
			_WhileOrUntilLocation = whileOrUntilLocation;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, Expression);
		}
	}
}

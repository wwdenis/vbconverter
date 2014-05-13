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
using VBConverter.CodeParser.Trees.Expressions;

namespace VBConverter.CodeParser.Trees.Initializers
{
	/// <summary>
	/// A parse tree for an expression initializer.
	/// </summary>
	public sealed class ExpressionInitializer : Initializer
	{

		private readonly Expression _Expression;

		/// <summary>
		/// The expression.
		/// </summary>
		public Expression Expression {
			get { return _Expression; }
		}

		/// <summary>
		/// Constructs a new expression initializer parse tree.
		/// </summary>
		/// <param name="expression">The expression.</param>
		/// <param name="span">The location of the parse tree.</param>
		public ExpressionInitializer(Expression expression, Span span) : base(TreeType.ExpressionInitializer, span)
		{

			if (expression == null)
			{
				throw new ArgumentNullException("expression");
			}

			SetParent(expression);

			_Expression = expression;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, Expression);
		}
	}
}

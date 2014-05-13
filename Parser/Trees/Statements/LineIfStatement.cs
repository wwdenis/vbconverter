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
	/// A parse tree for a line If statement.
	/// </summary>
	public sealed class LineIfStatement : Statement
	{

		private readonly Expression _Expression;
		private readonly Location _ThenLocation;
		private readonly StatementCollection _IfStatements;
		private readonly Location _ElseLocation;
		private readonly StatementCollection _ElseStatements;

		/// <summary>
		/// The conditional expression.
		/// </summary>
		public Expression Expression {
			get { return _Expression; }
		}

		/// <summary>
		/// The location of the 'Then'.
		/// </summary>
		public Location ThenLocation {
			get { return _ThenLocation; }
		}

		/// <summary>
		/// The If statements.
		/// </summary>
		public StatementCollection IfStatements {
			get { return _IfStatements; }
		}

		/// <summary>
		/// The location of the 'Else', if any.
		/// </summary>
		public Location ElseLocation {
			get { return _ElseLocation; }
		}

		/// <summary>
		/// The Else statements.
		/// </summary>
		public StatementCollection ElseStatements {
			get { return _ElseStatements; }
		}

		/// <summary>
		/// Constructs a new parse tree for a line If statement.
		/// </summary>
		/// <param name="expression">The conditional expression.</param>
		/// <param name="thenLocation">The location of the 'Then'.</param>
		/// <param name="ifStatements">The If statements.</param>
		/// <param name="elseLocation">The location of the 'Else', if any.</param>
		/// <param name="elseStatements">The Else statements.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public LineIfStatement(Expression expression, Location thenLocation, StatementCollection ifStatements, Location elseLocation, StatementCollection elseStatements, Span span, IList<Comment> comments) : base(TreeType.LineIfBlockStatement, span, comments)
		{

			if (expression == null)
			{
				throw new ArgumentNullException("expression");
			}

			SetParent(expression);
			SetParent(ifStatements);
			SetParent(elseStatements);

			_Expression = expression;
			_ThenLocation = thenLocation;
			_IfStatements = ifStatements;
			_ElseLocation = elseLocation;
			_ElseStatements = elseStatements;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, Expression);
			AddChild(childList, IfStatements);
			AddChild(childList, ElseStatements);
		}
	}
}

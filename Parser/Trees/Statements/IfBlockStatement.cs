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
	/// A parse tree for an If block.
	/// </summary>
	public sealed class IfBlockStatement : BlockStatement
	{

		private readonly Expression _Expression;
		private readonly Location _ThenLocation;
		private readonly StatementCollection _ElseIfBlockStatements;
		private readonly ElseBlockStatement _ElseBlockStatement;
		private readonly EndBlockStatement _EndStatement;

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
		/// The Else If statements.
		/// </summary>
		public StatementCollection ElseIfBlockStatements {
			get { return _ElseIfBlockStatements; }
		}

		/// <summary>
		/// The Else statement, if any.
		/// </summary>
		public ElseBlockStatement ElseBlockStatement {
			get { return _ElseBlockStatement; }
		}

		/// <summary>
		/// The End If statement, if any.
		/// </summary>
		public EndBlockStatement EndStatement {
			get { return _EndStatement; }
		}

		/// <summary>
		/// Constructs a new parse tree for a If statement.
		/// </summary>
		/// <param name="expression">The conditional expression.</param>
		/// <param name="thenLocation">The location of the 'Then', if any.</param>
		/// <param name="statements">The statements in the If block.</param>
		/// <param name="elseIfBlockStatements">The Else If statements.</param>
		/// <param name="elseBlockStatement">The Else statement, if any.</param>
		/// <param name="endStatement">The End If statement, if any.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public IfBlockStatement(Expression expression, Location thenLocation, StatementCollection statements, StatementCollection elseIfBlockStatements, ElseBlockStatement elseBlockStatement, EndBlockStatement endStatement, Span span, IList<Comment> comments) : base(TreeType.IfBlockStatement, statements, span, comments)
		{

			if (expression == null)
			{
				throw new ArgumentNullException("expression");
			}

			SetParent(expression);
			SetParent(elseIfBlockStatements);
			SetParent(elseBlockStatement);
			SetParent(endStatement);

			_Expression = expression;
			_ThenLocation = thenLocation;
			_ElseIfBlockStatements = elseIfBlockStatements;
			_ElseBlockStatement = elseBlockStatement;
			_EndStatement = endStatement;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, Expression);
			base.GetChildTrees(childList);
			AddChild(childList, ElseIfBlockStatements);
			AddChild(childList, ElseBlockStatement);
			AddChild(childList, EndStatement);
		}
	}
}

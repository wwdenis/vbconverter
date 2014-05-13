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
	/// A parse tree for a Select statement.
	/// </summary>
	public sealed class SelectBlockStatement : BlockStatement
	{

		private readonly Location _CaseLocation;
		private readonly Expression _Expression;
		private readonly StatementCollection _CaseBlockStatements;
		private readonly CaseElseBlockStatement _CaseElseBlockStatement;
		private readonly EndBlockStatement _EndStatement;

		/// <summary>
		/// The location of the 'Case', if any.
		/// </summary>
		public Location CaseLocation {
			get { return _CaseLocation; }
		}

		/// <summary>
		/// The location of the select expression.
		/// </summary>
		public Expression Expression {
			get { return _Expression; }
		}

		/// <summary>
		/// The Case statements.
		/// </summary>
		public StatementCollection CaseBlockStatements {
			get { return _CaseBlockStatements; }
		}

		/// <summary>
		/// The Case Else statement, if any.
		/// </summary>
		public CaseElseBlockStatement CaseElseBlockStatement {
			get { return _CaseElseBlockStatement; }
		}

		/// <summary>
		/// The End Select statement, if any.
		/// </summary>
		public EndBlockStatement EndStatement {
			get { return _EndStatement; }
		}

		/// <summary>
		/// Constructs a new parse tree for a Select statement.
		/// </summary>
		/// <param name="caseLocation">The location of the 'Case', if any.</param>
		/// <param name="expression">The select expression.</param>
		/// <param name="caseBlockStatements">The Case statements.</param>
		/// <param name="caseElseBlockStatement">The Case Else statement, if any.</param>
		/// <param name="endStatement">The End Select statement, if any.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public SelectBlockStatement(Location caseLocation, Expression expression, StatementCollection statements, StatementCollection caseBlockStatements, CaseElseBlockStatement caseElseBlockStatement, EndBlockStatement endStatement, Span span, IList<Comment> comments) : base(TreeType.SelectBlockStatement, statements, span, comments)
		{

			if (expression == null)
			{
				throw new ArgumentNullException("expression");
			}

			SetParent(expression);
			SetParent(caseBlockStatements);
			SetParent(caseElseBlockStatement);
			SetParent(endStatement);

			_CaseLocation = caseLocation;
			_Expression = expression;
			_CaseBlockStatements = caseBlockStatements;
			_CaseElseBlockStatement = caseElseBlockStatement;
			_EndStatement = endStatement;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, Expression);
			AddChild(childList, CaseBlockStatements);
			AddChild(childList, CaseElseBlockStatement);
			AddChild(childList, EndStatement);
		}
	}
}

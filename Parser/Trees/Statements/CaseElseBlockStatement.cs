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
	/// A parse tree for the block of a Case Else statement.
	/// </summary>
	public sealed class CaseElseBlockStatement : BlockStatement
	{

		private readonly CaseElseStatement _CaseElseStatement;

		/// <summary>
		/// The Case Else statement that started the block.
		/// </summary>
		public CaseElseStatement CaseElseStatement {
			get { return _CaseElseStatement; }
		}

		/// <summary>
		/// Constructs a new parse tree for the block of a Case Else statement.
		/// </summary>
		/// <param name="caseElseStatement">The Case Else statement that started the block.</param>
		/// <param name="statements">The statements in the block.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments of the tree.</param>
		public CaseElseBlockStatement(CaseElseStatement caseElseStatement, StatementCollection statements, Span span, IList<Comment> comments) : base(TreeType.CaseElseBlockStatement, statements, span, comments)
		{

			if (caseElseStatement == null)
			{
				throw new ArgumentNullException("caseElseStatement");
			}

			SetParent(caseElseStatement);
			_CaseElseStatement = caseElseStatement;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, CaseElseStatement);
			base.GetChildTrees(childList);
		}
	}
}

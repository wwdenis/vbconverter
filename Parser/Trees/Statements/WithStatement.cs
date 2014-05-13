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
	/// A parse tree for a With block statement.
	/// </summary>
	public sealed class WithBlockStatement : ExpressionBlockStatement
	{
        public StatementCollection QualifiedStatements
        {
            get
            {
                StatementCollection statements = null;

                if (Statements != null)
                {
                    statements = new StatementCollection(Statements, Statements.ColonLocations, Statements.Span);
                    SetParent(statements);
                    foreach (Statement item in statements)
                        FillQualifiedStatement(item);
                }

                return statements;
            }
        }

        private void FillQualifiedStatement(Tree tree)
        {
            if (tree == null)
                return;

            if (tree is QualifiedExpression)
            {
                QualifiedExpression qualified = (QualifiedExpression)tree;
                if (qualified.Qualifier == null)
                    qualified.Qualifier = Expression;
            }

            if (tree.Children == null)
                return;

            foreach (Tree item in tree.Children)
                FillQualifiedStatement(item);
        }

        /// <summary>
		/// Constructs a new parse tree for a With statement block.
		/// </summary>
		/// <param name="expression">The expression.</param>
		/// <param name="statements">The statements in the block.</param>
		/// <param name="endStatement">The End statement for the block, if any.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public WithBlockStatement(Expression expression, StatementCollection statements, EndBlockStatement endStatement, Span span, IList<Comment> comments) : base(TreeType.WithBlockStatement, expression, statements, endStatement, span, comments)
		{
		}
	}
}

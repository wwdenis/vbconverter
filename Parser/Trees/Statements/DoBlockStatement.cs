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
    /// A parse tree for a Do statement.
    /// </summary>
    public sealed class DoBlockStatement : BlockStatement
    {

        private readonly bool _IsWhile;
        private readonly Location _WhileOrUntilLocation;
        private readonly Expression _Expression;
        private readonly LoopStatement _EndStatement;

        /// <summary>
        /// Whether the Do is followed by a While or Until, if any.
        /// </summary>
        public bool IsWhile
        {
            get { return _IsWhile; }
        }

        /// <summary>
        /// The location of the While or Until, if any.
        /// </summary>
        public Location WhileOrUntilLocation
        {
            get { return _WhileOrUntilLocation; }
        }

        /// <summary>
        /// The While or Until expression, if any.
        /// </summary>
        public Expression Expression
        {
            get { return _Expression; }
        }

        /// <summary>
        /// The ending Loop statement.
        /// </summary>
        public LoopStatement EndStatement
        {
            get { return _EndStatement; }
        }

        /// <summary>
        /// Constructs a new parse tree for a Do statement.
        /// </summary>
        /// <param name="expression">The While or Until expression, if any.</param>
        /// <param name="isWhile">Whether the Do is followed by a While or Until, if any.</param>
        /// <param name="whileOrUntilLocation">The location of the While or Until, if any.</param>
        /// <param name="statements">The statements in the block.</param>
        /// <param name="endStatement">The ending Loop statement.</param>
        /// <param name="span">The location of the parse tree.</param>
        /// <param name="comments">The comments on the parse tree.</param>
        public DoBlockStatement(Expression expression, bool isWhile, Location whileOrUntilLocation, StatementCollection statements, LoopStatement endStatement, Span span, IList<Comment> comments) : base(TreeType.DoBlockStatement, statements, span, comments)
        {

            SetParent(expression);
            SetParent(endStatement);

            _Expression = expression;
            _IsWhile = isWhile;
            _WhileOrUntilLocation = whileOrUntilLocation;
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

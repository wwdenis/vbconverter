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
using VBConverter.CodeParser.Trees.Statements;
using VBConverter.CodeParser.Trees.Types;

namespace VBConverter.CodeParser.Trees.CaseClause
{
    /// <summary>
    /// A parse tree for a case clause that compares against a range of values.
    /// </summary>
    public sealed class RangeCaseClause : CaseClause
    {
        private readonly Expression _RangeExpression;

        /// <summary>
        /// The range expression.
        /// </summary>
        public Expression RangeExpression
        {
            get { return _RangeExpression; }
        }

        public override BinaryOperatorExpression ComparisonExpression
        {
            get 
            {
                if (SelectCaseBlock == null || RangeExpression == null)
                    return null;

                Expression expression = SelectCaseBlock.Expression;
                BinaryOperatorExpression result = null;
                
                if (RangeExpression is BinaryOperatorExpression)
                {
                    Expression left = ((BinaryOperatorExpression)RangeExpression).LeftOperand;
                    Expression right = ((BinaryOperatorExpression)RangeExpression).RightOperand;

                    BinaryOperatorExpression minor = new BinaryOperatorExpression(expression, BinaryOperatorType.GreaterThanEquals, left);
                    BinaryOperatorExpression major = new BinaryOperatorExpression(expression, BinaryOperatorType.LessThanEquals, right);

                    result = new BinaryOperatorExpression(minor, BinaryOperatorType.And, major);
                }
                else
                {
                    result = new BinaryOperatorExpression(expression, BinaryOperatorType.Equals, RangeExpression);
                }

                return result;
            }
        }

        /// <summary>
        /// Constructs a new range case clause parse tree.
        /// </summary>
        /// <param name="rangeExpression">The range expression.</param>
        /// <param name="span">The location of the parse tree.</param>
        public RangeCaseClause(Expression rangeExpression, Span span) : base(TreeType.RangeCaseClause, span)
        {
            if (rangeExpression == null)
                throw new ArgumentNullException("rangeExpression");

            SetParent(rangeExpression);

            _RangeExpression = rangeExpression;
        }

        protected override void GetChildTrees(IList<Tree> childList)
        {
            AddChild(childList, RangeExpression);
        }
    }
}

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
using VBConverter.CodeParser.Trees.Types;

namespace VBConverter.CodeParser.Trees.Statements
{
	/// <summary>
	/// A parse tree for a compound assignment statement.
	/// </summary>
	public sealed class CompoundAssignmentStatement : Statement
	{

		private readonly Expression _TargetExpression;
		private readonly BinaryOperatorType _CompoundOperator;
		private readonly Location _OperatorLocation;
		private readonly Expression _SourceExpression;

		/// <summary>
		/// The target of the assignment.
		/// </summary>
		public Expression TargetExpression {
			get { return _TargetExpression; }
		}

		/// <summary>
		/// The compound operator.
		/// </summary>
		public BinaryOperatorType CompoundOperator {
			get { return _CompoundOperator; }
		}

		/// <summary>
		/// The location of the operator.
		/// </summary>
		public Location OperatorLocation {
			get { return _OperatorLocation; }
		}

		/// <summary>
		/// The source of the assignment.
		/// </summary>
		public Expression SourceExpression {
			get { return _SourceExpression; }
		}

		/// <summary>
		/// Constructs a new parse tree for a compound assignment statement.
		/// </summary>
		/// <param name="compoundOperator">The compound operator.</param>
		/// <param name="targetExpression">The target of the assignment.</param>
		/// <param name="operatorLocation">The location of the operator.</param>
		/// <param name="sourceExpression">The source of the assignment.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public CompoundAssignmentStatement(BinaryOperatorType compoundOperator, Expression targetExpression, Location operatorLocation, Expression sourceExpression, Span span, IList<Comment> comments) : base(TreeType.CompoundAssignmentStatement, span, comments)
		{

			if (compoundOperator < BinaryOperatorType.Plus || compoundOperator > BinaryOperatorType.Power)
				throw new ArgumentOutOfRangeException("compoundOperator");
			
			if (targetExpression == null)
				throw new ArgumentNullException("targetExpression");
			
			if (sourceExpression == null)
				throw new ArgumentNullException("sourceExpression");
		
			SetParent(targetExpression);
			SetParent(sourceExpression);

			_CompoundOperator = compoundOperator;
			_TargetExpression = targetExpression;
			_OperatorLocation = operatorLocation;
			_SourceExpression = sourceExpression;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, TargetExpression);
			AddChild(childList, SourceExpression);
		}
	}
}

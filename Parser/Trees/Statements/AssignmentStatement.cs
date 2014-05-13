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
using VBConverter.CodeParser.Tokens;
using VBConverter.CodeParser.Trees.Comments;
using VBConverter.CodeParser.Trees.Expressions;

namespace VBConverter.CodeParser.Trees.Statements
{
	/// <summary>
	/// A parse tree for an assignment statement.
	/// </summary>
	public sealed class AssignmentStatement : Statement
	{

		private readonly Expression _TargetExpression;
		private readonly Location _OperatorLocation;
		private readonly Expression _SourceExpression;
        private readonly Token _Acessor;

        /// <summary>
        /// The assignment acessor
        /// </summary>
        public Token Acessor {
            get { return _Acessor; }
        }

        /// <summary>
        /// The type of the assignment acessor
        /// </summary>
        public TokenType AcessorType {
            get { return (_Acessor == null ? TokenType.None : _Acessor.Type); }
        }

        /// <summary>
        /// The location of the assignment acessor
        /// </summary>
        public Location AcessorLocation {
            get { return (_Acessor == null ? Location.Empty : _Acessor.Span.Start); }
        }

		/// <summary>
		/// The target of the assignment.
		/// </summary>
		public Expression TargetExpression {
			get { return _TargetExpression; }
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
		/// Constructs a new parse tree for an assignment statement.
		/// </summary>
		/// <param name="targetExpression">The target of the assignment.</param>
		/// <param name="operatorLocation">The location of the operator.</param>
		/// <param name="sourceExpression">The source of the assignment.</param>
		/// <param name="span">The location of the parse tree.</param>
        /// <param name="acessor">The acessor token of the assignment.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public AssignmentStatement(Expression targetExpression, Location operatorLocation, Expression sourceExpression, Span span, IList<Comment> comments, Token acessor) 
            : base(TreeType.AssignmentStatement, span, comments)
		{
			if (targetExpression == null)
				throw new ArgumentNullException("targetExpression");
		
			if (sourceExpression == null)
				throw new ArgumentNullException("sourceExpression");
		
			SetParent(targetExpression);
			SetParent(sourceExpression);

			_TargetExpression = targetExpression;
			_OperatorLocation = operatorLocation;
			_SourceExpression = sourceExpression;
            _Acessor = acessor;
		}

        /// <summary>
		/// Constructs a new parse tree for an assignment statement.
		/// </summary>
		/// <param name="targetExpression">The target of the assignment.</param>
		/// <param name="operatorLocation">The location of the operator.</param>
		/// <param name="sourceExpression">The source of the assignment.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
        public AssignmentStatement(Expression targetExpression, Location operatorLocation, Expression sourceExpression, Span span, IList<Comment> comments)
            : this(targetExpression, operatorLocation, sourceExpression, span, comments, null)
        {
        }

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, TargetExpression);
			AddChild(childList, SourceExpression);
		}
	}
}

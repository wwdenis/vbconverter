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
	/// A parse tree for a Mid assignment statement.
	/// </summary>
	public sealed class MidAssignmentStatement : Statement
	{

		private readonly bool _HasTypeCharacter;
		private readonly Location _LeftParenthesisLocation;
		private readonly Expression _TargetExpression;
		private readonly Location _StartCommaLocation;
		private readonly Expression _StartExpression;
		private readonly Location _LengthCommaLocation;
		private readonly Expression _LengthExpression;
		private readonly Location _RightParenthesisLocation;
		private readonly Location _OperatorLocation;
		private readonly Expression _SourceExpression;

		/// <summary>
		/// Whether the Mid identifier had a type character.
		/// </summary>
		public bool HasTypeCharacter {
			get { return _HasTypeCharacter; }
		}

		/// <summary>
		/// The location of the left parenthesis.
		/// </summary>
		public Location LeftParenthesisLocation {
			get { return _LeftParenthesisLocation; }
		}

		/// <summary>
		/// The target of the assignment.
		/// </summary>
		public Expression TargetExpression {
			get { return _TargetExpression; }
		}

		/// <summary>
		/// The location of the comma before the start expression.
		/// </summary>
		public Location StartCommaLocation {
			get { return _StartCommaLocation; }
		}

		/// <summary>
		/// The expression representing the start of the string to replace.
		/// </summary>
		public Expression StartExpression {
			get { return _StartExpression; }
		}

		/// <summary>
		/// The location of the comma before the length expression, if any.
		/// </summary>
		public Location LengthCommaLocation {
			get { return _LengthCommaLocation; }
		}

		/// <summary>
		/// The expression representing the length of the string to replace, if any.
		/// </summary>
		public Expression LengthExpression {
			get { return _LengthExpression; }
		}

		/// <summary>
		/// The right parenthesis location.
		/// </summary>
		public Location RightParenthesisLocation {
			get { return _RightParenthesisLocation; }
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
		/// <param name="hasTypeCharacter">Whether the Mid identifier has a type character.</param>
		/// <param name="leftParenthesisLocation">The location of the left parenthesis.</param>
		/// <param name="targetExpression">The target of the assignment.</param>
		/// <param name="startCommaLocation">The location of the comma before the start expression.</param>
		/// <param name="startExpression">The expression representing the start of the string to replace.</param>
		/// <param name="lengthCommaLocation">The location of the comma before the length expression, if any.</param>
		/// <param name="lengthExpression">The expression representing the length of the string to replace, if any.</param>
		/// <param name="operatorLocation">The location of the operator.</param>
		/// <param name="sourceExpression">The source of the assignment.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public MidAssignmentStatement(bool hasTypeCharacter, Location leftParenthesisLocation, Expression targetExpression, Location startCommaLocation, Expression startExpression, Location lengthCommaLocation, Expression lengthExpression, Location rightParenthesisLocation, Location operatorLocation, Expression sourceExpression, 
		Span span, IList<Comment> comments) : base(TreeType.MidAssignmentStatement, span, comments)
		{

			if (targetExpression == null)
			{
				throw new ArgumentNullException("targetExpression");
			}

			if (startExpression == null)
			{
				throw new ArgumentNullException("startExpression");
			}

			if (sourceExpression == null)
			{
				throw new ArgumentNullException("sourceExpression");
			}

			SetParent(targetExpression);
			SetParent(startExpression);
			SetParent(lengthExpression);
			SetParent(sourceExpression);

			_HasTypeCharacter = hasTypeCharacter;
			_LeftParenthesisLocation = leftParenthesisLocation;
			_TargetExpression = targetExpression;
			_StartCommaLocation = startCommaLocation;
			_StartExpression = startExpression;
			_LengthCommaLocation = lengthCommaLocation;
			_LengthExpression = lengthExpression;
			_RightParenthesisLocation = rightParenthesisLocation;
			_OperatorLocation = operatorLocation;
			_SourceExpression = sourceExpression;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, TargetExpression);
			AddChild(childList, StartExpression);
			AddChild(childList, LengthExpression);
			AddChild(childList, SourceExpression);
		}
	}
}

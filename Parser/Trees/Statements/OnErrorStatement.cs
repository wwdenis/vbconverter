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
using VBConverter.CodeParser.Trees.Names;

namespace VBConverter.CodeParser.Trees.Statements
{
	/// <summary>
	/// A parse tree for an On Error statement.
	/// </summary>
	public sealed class OnErrorStatement : LabelReferenceStatement
	{

		private readonly OnErrorType _OnErrorType;
		private readonly Location _ErrorLocation;
		private readonly Location _ResumeOrGoToLocation;
		private readonly Location _NextOrZeroOrMinusLocation;
		private readonly Location _OneLocation;

		/// <summary>
		/// The type of On Error statement.
		/// </summary>
		public OnErrorType OnErrorType {
			get { return _OnErrorType; }
		}

		/// <summary>
		/// The location of the 'Error'.
		/// </summary>
		public Location ErrorLocation {
			get { return _ErrorLocation; }
		}

		/// <summary>
		/// The location of the 'Resume' or 'GoTo'.
		/// </summary>
		public Location ResumeOrGoToLocation {
			get { return _ResumeOrGoToLocation; }
		}

		/// <summary>
		/// The location of the 'Next', '0' or '-', if any.
		/// </summary>
		public Location NextOrZeroOrMinusLocation {
			get { return _NextOrZeroOrMinusLocation; }
		}

		/// <summary>
		/// The location of the '1', if any.
		/// </summary>
		public Location OneLocation {
			get { return _OneLocation; }
		}

		/// <summary>
		/// Constructs a parse tree for an On Error statement.
		/// </summary>
		/// <param name="onErrorType">The type of the On Error statement.</param>
		/// <param name="errorLocation">The location of the 'Error'.</param>
		/// <param name="resumeOrGoToLocation">The location of the 'Resume' or 'GoTo'.</param>
		/// <param name="nextOrZeroOrMinusLocation">The location of the 'Next', '0' or '-', if any.</param>
		/// <param name="oneLocation">The location of the '1', if any.</param>
		/// <param name="name">The label to branch to, if any.</param>
		/// <param name="isLineNumber">Whether the label is a line number.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public OnErrorStatement(OnErrorType onErrorType, Location errorLocation, Location resumeOrGoToLocation, Location nextOrZeroOrMinusLocation, Location oneLocation, SimpleName name, bool isLineNumber, Span span, IList<Comment> comments) : base(TreeType.OnErrorStatement, name, isLineNumber, span, comments)
		{

            if (!Enum.IsDefined(typeof(OnErrorType), onErrorType))
				throw new ArgumentOutOfRangeException("onErrorType");
			
			_OnErrorType = onErrorType;
			_ErrorLocation = errorLocation;
			_ResumeOrGoToLocation = resumeOrGoToLocation;
			_NextOrZeroOrMinusLocation = nextOrZeroOrMinusLocation;
			_OneLocation = oneLocation;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			if (OnErrorType == OnErrorType.Label)
			{
				base.GetChildTrees(childList);
			}
		}
	}
}

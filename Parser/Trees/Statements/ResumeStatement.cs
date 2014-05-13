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
	/// A parse tree for a Resume statement.
	/// </summary>
	public sealed class ResumeStatement : LabelReferenceStatement
	{

		private readonly ResumeType _ResumeType;
		private readonly Location _NextLocation;

		/// <summary>
		/// The type of the Resume statement.
		/// </summary>
		public ResumeType ResumeType {
			get { return _ResumeType; }
		}

		/// <summary>
		/// The location of the 'Next', if any.
		/// </summary>
		public Location NextLocation {
			get { return _NextLocation; }
		}

		/// <summary>
		/// Constructs a parse tree for a Resume statement.
		/// </summary>
		/// <param name="resumeType">The type of the Resume statement.</param>
		/// <param name="nextLocation">The location of the 'Next', if any.</param>
		/// <param name="name">The label name, if any.</param>
		/// <param name="isLineNumber">Whether the label is a line number.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments of the parse tree.</param>
		public ResumeStatement(ResumeType resumeType, Location nextLocation, SimpleName name, bool isLineNumber, Span span, IList<Comment> comments) : base(TreeType.ResumeStatement, name, isLineNumber, span, comments)
		{

            if (!Enum.IsDefined(typeof(ResumeType), resumeType))
				throw new ArgumentOutOfRangeException("resumeType");
			
			_ResumeType = resumeType;
			_NextLocation = nextLocation;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			if (ResumeType == ResumeType.Label)
			{
				base.GetChildTrees(childList);
			}
		}
	}
}

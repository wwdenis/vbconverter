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
using VBConverter.CodeParser.Trees.Types;

namespace VBConverter.CodeParser.Trees.Comments
{
	/// <summary>
	/// A parse tree for a comment.
	/// </summary>
	public sealed class Comment : Tree
	{

		private readonly string _Text;
		private readonly bool _IsREM;

		/// <summary>
		/// The text of the comment.
		/// </summary>
		public string Text {
			get { return _Text; }
		}

		/// <summary>
		/// Whether the comment is a REM comment.
		/// </summary>
		public bool IsREM {
			get { return _IsREM; }
		}

		/// <summary>
		/// Constructs a new comment parse tree.
		/// </summary>
		/// <param name="comment">The text of the comment.</param>
		/// <param name="isREM">Whether the comment is a REM comment.</param>
		/// <param name="span">The location of the parse tree.</param>
		public Comment(string comment, bool isREM, Span span) : base(TreeType.Comment, span)
		{

			if (comment == null)
			{
				throw new ArgumentNullException("comment");
			}

			_Text = comment;
			_IsREM = isREM;
		}
	}
}

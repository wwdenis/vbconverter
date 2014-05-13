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
namespace VBConverter.CodeParser.Tokens.Terminator
{
	/// <summary>
	/// A comment token.
	/// </summary>
	public sealed class CommentToken : Token
	{
		// Was the comment preceded by a quote or by REM?
		private readonly bool _IsREM;
		// Comment can be Nothing
		private readonly string _Comment;

		/// <summary>
		/// Whether the comment was preceded by REM.
		/// </summary>
		public bool IsREM {
			get { return _IsREM; }
		}

		/// <summary>
		/// The text of the comment.
		/// </summary>
		public string Comment {
			get { return _Comment; }
		}

		/// <summary>
		/// Constructs a new comment token.
		/// </summary>
		/// <param name="comment">The comment value.</param>
		/// <param name="isREM">Whether the comment was preceded by REM.</param>
		/// <param name="span">The location of the comment.</param>
		public CommentToken(string comment, bool isREM, Span span) : base(TokenType.Comment, span)
		{
			if (comment == null)
				throw new ArgumentNullException("comment");

			_IsREM = isREM;
			_Comment = comment;
		}
	}
}

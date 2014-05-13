using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
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
using VBConverter.CodeParser.Trees.Types;

namespace VBConverter.CodeParser.Trees.Declarations
{
	/// <summary>
	/// A parse tree for a declaration.
	/// </summary>
	public abstract class Declaration : CommentableTree
	{
        public override bool IsBad
        {
            get { return Type == TreeType.SyntaxError; }
        }

		/// <summary>
		/// Creates a bad declaration.
		/// </summary>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		/// <returns>A bad declaration.</returns>
		public static Declaration GetBadDeclaration(Span span, IList<Comment> comments)
		{
			return new BadDeclaration(span, comments);
		}

		protected Declaration(TreeType type, Span span, IList<Comment> comments) : base(type, span, comments)
		{
            if (type != TreeType.SyntaxError && !Tree.Types.IsDeclaration(type))
                throw new ArgumentOutOfRangeException("type", "Invalid declaration type: " + type.ToString());
		}
	}
}

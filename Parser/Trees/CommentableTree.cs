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

namespace VBConverter.CodeParser.Trees
{
	/// <summary>
	/// A parse tree for a declaration.
	/// </summary>
	public abstract class CommentableTree : Tree
	{
		private readonly ReadOnlyCollection<Comment> _Comments;

		/// <summary>
		/// The comments for the tree.
		/// </summary>
		public ReadOnlyCollection<Comment> Comments
        {
			get { return _Comments; }
		}

        protected CommentableTree(TreeType type, Span span, IList<Comment> comments) : base(type, span)
		{
			if (comments != null)
				_Comments = new ReadOnlyCollection<Comment>(comments);
		}
	}
}

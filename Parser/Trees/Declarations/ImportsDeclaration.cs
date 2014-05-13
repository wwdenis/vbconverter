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
using VBConverter.CodeParser.Trees.Imports;

namespace VBConverter.CodeParser.Trees.Declarations
{
	/// <summary>
	/// A parse tree for an Imports declaration.
	/// </summary>
	public sealed class ImportsDeclaration : Declaration
	{

		private readonly ImportCollection _ImportMembers;

		/// <summary>
		/// The members imported.
		/// </summary>
		public ImportCollection ImportMembers {
			get { return _ImportMembers; }
		}

		/// <summary>
		/// Constructs a parse tree for an Imports declaration.
		/// </summary>
		/// <param name="importMembers">The members imported.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public ImportsDeclaration(ImportCollection importMembers, Span span, IList<Comment> comments) : base(TreeType.ImportsDeclaration, span, comments)
		{

			if (importMembers == null)
			{
				throw new ArgumentNullException("importMembers");
			}

			SetParent(importMembers);

			_ImportMembers = importMembers;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			base.GetChildTrees(childList);

			AddChild(childList, ImportMembers);
		}
	}
}

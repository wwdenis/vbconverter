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
using VBConverter.CodeParser.Trees.Attributes;
using VBConverter.CodeParser.Trees.Comments;
using VBConverter.CodeParser.Trees.Types;

namespace VBConverter.CodeParser.Trees.Declarations
{
	/// <summary>
	/// A parse tree for an assembly-level or module-level attribute declaration.
	/// </summary>
	public sealed class AttributeDeclaration : Declaration
	{

		private readonly AttributeBlockCollection _Attributes;

		/// <summary>
		/// The attributes.
		/// </summary>
		public AttributeBlockCollection Attributes {
			get { return _Attributes; }
		}

		/// <summary>
		/// Constructs a new parse tree for assembly-level or module-level attribute declarations.
		/// </summary>
		/// <param name="attributes">The attributes.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public AttributeDeclaration(AttributeBlockCollection attributes, Span span, IList<Comment> comments) : base(TreeType.AttributeDeclaration, span, comments)
		{

			if (attributes == null)
			{
				throw new ArgumentNullException("attributes");
			}

			SetParent(attributes);

			_Attributes = attributes;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, Attributes);
		}
	}
}

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
using VBConverter.CodeParser.Trees.Declarations;
using VBConverter.CodeParser.Trees.Modifiers;
using VBConverter.CodeParser.Trees.Statements;

namespace VBConverter.CodeParser.Trees.Declarations.Members
{
	/// <summary>
	/// A parse tree for a Get property accessor.
	/// </summary>
	public sealed class GetAccessorDeclaration : ModifiedDeclaration
	{

		private readonly Location _GetLocation;
		private readonly StatementCollection _Statements;
		private readonly EndBlockDeclaration _EndDeclaration;

		/// <summary>
		/// The location of the 'Get'.
		/// </summary>
		public Location GetLocation {
			get { return _GetLocation; }
		}

		/// <summary>
		/// The statements in the accessor.
		/// </summary>
		public StatementCollection Statements {
			get { return _Statements; }
		}

		/// <summary>
		/// The End declaration for the accessor.
		/// </summary>
		public EndBlockDeclaration EndDeclaration {
			get { return _EndDeclaration; }
		}

		/// <summary>
		/// Constructs a new parse tree for a Get property accessor.
		/// </summary>
		/// <param name="attributes">The attributes for the parse tree.</param>
		/// <param name="modifiers">The modifiers for the parse tree.</param>
		/// <param name="getLocation">The location of the 'Get'.</param>
		/// <param name="statements">The statements in the declaration.</param>
		/// <param name="endDeclaration">The end block declaration, if any.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public GetAccessorDeclaration(AttributeBlockCollection attributes, ModifierCollection modifiers, Location getLocation, StatementCollection statements, EndBlockDeclaration endDeclaration, Span span, IList<Comment> comments) : base(TreeType.GetAccessorDeclaration, attributes, modifiers, span, comments)
		{

			SetParent(statements);
			SetParent(endDeclaration);

			_GetLocation = getLocation;
			_Statements = statements;
			_EndDeclaration = endDeclaration;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			base.GetChildTrees(childList);

			AddChild(childList, Statements);
			AddChild(childList, EndDeclaration);
		}
	}
}

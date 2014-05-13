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
using VBConverter.CodeParser.Trees.Names;
using VBConverter.CodeParser.Trees.TypeNames;

namespace VBConverter.CodeParser.Trees.Declarations.Members
{
	/// <summary>
	/// A parse tree for a custom event declaration.
	/// </summary>
	public sealed class CustomEventDeclaration : SignatureDeclaration
	{

		private readonly Location _CustomLocation;
		private readonly NameCollection _ImplementsList;
		private readonly DeclarationCollection _Accessors;
		private readonly EndBlockDeclaration _EndDeclaration;

		/// <summary>
		/// The location of the 'Custom' keyword.
		/// </summary>
		public Location CustomLocation {
			get { return _CustomLocation; }
		}

		/// <summary>
		/// The list of implemented members.
		/// </summary>
		public NameCollection ImplementsList {
			get { return _ImplementsList; }
		}

		/// <summary>
		/// The event accessors.
		/// </summary>
		public DeclarationCollection Accessors {
			get { return _Accessors; }
		}

		/// <summary>
		/// The End Event declaration, if any.
		/// </summary>
		public EndBlockDeclaration EndDeclaration {
			get { return _EndDeclaration; }
		}

		/// <summary>
		/// Constructs a new parse tree for a custom property declaration.
		/// </summary>
		/// <param name="attributes">The attributes on the declaration.</param>
		/// <param name="modifiers">The modifiers on the declaration.</param>
		/// <param name="customLocation">The location of the 'Custom' keyword.</param>
		/// <param name="keywordLocation">The location of the keyword.</param>
		/// <param name="name">The name of the custom event.</param>
		/// <param name="asLocation">The location of the 'As', if any.</param>
		/// <param name="resultType">The result type, if any.</param>
		/// <param name="implementsList">The implements list.</param>
		/// <param name="accessors">The custom event accessors.</param>
		/// <param name="endDeclaration">The End Event declaration, if any.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public CustomEventDeclaration(AttributeBlockCollection attributes, ModifierCollection modifiers, Location customLocation, Location keywordLocation, SimpleName name, Location asLocation, TypeName resultType, NameCollection implementsList, DeclarationCollection accessors, EndBlockDeclaration endDeclaration, 
		Span span, IList<Comment> comments) : base(TreeType.CustomEventDeclaration, attributes, modifiers, keywordLocation, name, null, null, asLocation, null, resultType, 
		span, comments)
		{

			SetParent(accessors);
			SetParent(endDeclaration);
			SetParent(implementsList);

			_CustomLocation = customLocation;
			_ImplementsList = implementsList;
			_Accessors = accessors;
			_EndDeclaration = endDeclaration;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			base.GetChildTrees(childList);
			AddChild(childList, ImplementsList);
			AddChild(childList, Accessors);
			AddChild(childList, EndDeclaration);
		}
	}
}

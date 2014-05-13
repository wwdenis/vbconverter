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
using VBConverter.CodeParser.Trees.Modifiers;
using VBConverter.CodeParser.Trees.Names;
using VBConverter.CodeParser.Trees.Parameters;
using VBConverter.CodeParser.Trees.TypeNames;
using VBConverter.CodeParser.Trees.TypeParameters;

namespace VBConverter.CodeParser.Trees.Declarations
{
	/// <summary>
	/// A parse tree for a delegate Function declaration.
	/// </summary>
	public sealed class DelegateFunctionDeclaration : DelegateDeclaration
	{

		/// <summary>
		/// Constructs a new parse tree for a delegate declaration.
		/// </summary>
		/// <param name="attributes">The attributes for the parse tree.</param>
		/// <param name="modifiers">The modifiers for the parse tree.</param>
		/// <param name="keywordLocation">The location of the keyword.</param>
		/// <param name="functionLocation">The location of the 'Function'.</param>
		/// <param name="name">The name of the declaration.</param>
		/// <param name="typeParameters">The type parameters of the declaration, if any.</param>
		/// <param name="parameters">The parameters of the declaration.</param>
		/// <param name="asLocation">The location of the 'As', if any.</param>
		/// <param name="resultTypeAttributes">The attributes on the result type, if any.</param>
		/// <param name="resultType">The result type, if any.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public DelegateFunctionDeclaration(AttributeBlockCollection attributes, ModifierCollection modifiers, Location keywordLocation, Location functionLocation, SimpleName name, TypeParameterCollection typeParameters, ParameterCollection parameters, Location asLocation, AttributeBlockCollection resultTypeAttributes, TypeName resultType, 
		Span span, IList<Comment> comments) : base(TreeType.DelegateFunctionDeclaration, attributes, modifiers, keywordLocation, functionLocation, name, typeParameters, parameters, asLocation, resultTypeAttributes, 
		resultType, span, comments)
		{
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			base.GetChildTrees(childList);
		}
	}
}

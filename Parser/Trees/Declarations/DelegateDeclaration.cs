using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
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
using VBConverter.CodeParser.Trees.Declarations.Members;
using VBConverter.CodeParser.Trees.Modifiers;
using VBConverter.CodeParser.Trees.Names;
using VBConverter.CodeParser.Trees.Parameters;
using VBConverter.CodeParser.Trees.TypeNames;
using VBConverter.CodeParser.Trees.TypeParameters;
using VBConverter.CodeParser.Trees.Types;

namespace VBConverter.CodeParser.Trees.Declarations
{
	/// <summary>
	/// A parse tree for a delegate declaration.
	/// </summary>
	public abstract class DelegateDeclaration : SignatureDeclaration
	{

		private readonly Location _SubOrFunctionLocation;

		/// <summary>
		/// The location of 'Sub' or 'Function'.
		/// </summary>
		public Location SubOrFunctionLocation {
			get { return _SubOrFunctionLocation; }
		}

		protected DelegateDeclaration(TreeType type, AttributeBlockCollection attributes, ModifierCollection modifiers, Location keywordLocation, Location subOrFunctionLocation, SimpleName name, TypeParameterCollection typeParameters, ParameterCollection parameters, Location asLocation, AttributeBlockCollection resultTypeAttributes, 
		TypeName resultType, Span span, IList<Comment> comments) : base(type, attributes, modifiers, keywordLocation, name, typeParameters, parameters, asLocation, resultTypeAttributes, resultType, 
		span, comments)
		{

			Debug.Assert(type == TreeType.DelegateSubDeclaration || type == TreeType.DelegateFunctionDeclaration);

			_SubOrFunctionLocation = subOrFunctionLocation;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			base.GetChildTrees(childList);
		}
	}
}

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
using VBConverter.CodeParser.Base;
using VBConverter.CodeParser.Trees.Attributes;
using VBConverter.CodeParser.Trees.Comments;
using VBConverter.CodeParser.Trees.Expressions;
using VBConverter.CodeParser.Trees.Parameters;
using VBConverter.CodeParser.Trees.Modifiers;
using VBConverter.CodeParser.Trees.Names;
using VBConverter.CodeParser.Trees.TypeNames;

namespace VBConverter.CodeParser.Trees.Declarations.Members
{
	/// <summary>
	/// A parse tree for a Declare Function statement.
	/// </summary>
	public sealed class ExternalFunctionDeclaration : ExternalDeclaration
	{

		/// <summary>
		/// Constructs a parse tree for a Declare Function statement.
		/// </summary>
		/// <param name="attributes">The attributes for the parse tree.</param>
		/// <param name="modifiers">The modifiers for the parse tree.</param>
		/// <param name="keywordLocation">The location of the keyword.</param>
		/// <param name="charsetLocation">The location of the 'Ansi', 'Auto' or 'Unicode', if any.</param>
		/// <param name="charset">The charset.</param>
		/// <param name="functionLocation">The location of 'Function'.</param>
		/// <param name="name">The name of the declaration.</param>
		/// <param name="libLocation">The location of 'Lib', if any.</param>
		/// <param name="libLiteral">The library, if any.</param>
		/// <param name="aliasLocation">The location of 'Alias', if any.</param>
		/// <param name="aliasLiteral">The alias, if any.</param>
		/// <param name="parameters">The parameters of the declaration.</param>
		/// <param name="asLocation">The location of the 'As', if any.</param>
		/// <param name="resultTypeAttributes">The attributes on the result type, if any.</param>
		/// <param name="resultType">The result type, if any.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public ExternalFunctionDeclaration(AttributeBlockCollection attributes, ModifierCollection modifiers, Location keywordLocation, Location charsetLocation, Charset charset, Location functionLocation, SimpleName name, Location libLocation, StringLiteralExpression libLiteral, Location aliasLocation, 
		StringLiteralExpression aliasLiteral, ParameterCollection parameters, Location asLocation, AttributeBlockCollection resultTypeAttributes, TypeName resultType, Span span, IList<Comment> comments) : base(TreeType.ExternalFunctionDeclaration, attributes, modifiers, keywordLocation, charsetLocation, charset, functionLocation, name, libLocation, libLiteral, 
		aliasLocation, aliasLiteral, parameters, asLocation, resultTypeAttributes, resultType, span, comments)
		{
		}
	}
}

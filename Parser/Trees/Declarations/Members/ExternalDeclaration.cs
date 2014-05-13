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
using VBConverter.CodeParser.Trees.Declarations;
using VBConverter.CodeParser.Trees.Expressions;
using VBConverter.CodeParser.Trees.Modifiers;
using VBConverter.CodeParser.Trees.Names;
using VBConverter.CodeParser.Trees.Parameters;
using VBConverter.CodeParser.Trees.TypeNames;

namespace VBConverter.CodeParser.Trees.Declarations.Members
{
	/// <summary>
	/// A parse tree for a Declare statement.
	/// </summary>
	public abstract class ExternalDeclaration : SignatureDeclaration
	{

		private readonly Location _CharsetLocation;
		private readonly Charset _Charset;
		private readonly Location _SubOrFunctionLocation;
		private readonly Location _LibLocation;
		private readonly StringLiteralExpression _LibLiteral;
		private readonly Location _AliasLocation;
		private readonly StringLiteralExpression _AliasLiteral;

		/// <summary>
		/// The location of 'Auto', 'Ansi' or 'Unicode', if any.
		/// </summary>
		public Location CharsetLocation {
			get { return _CharsetLocation; }
		}

		/// <summary>
		/// The charset.
		/// </summary>
		public Charset Charset {
			get { return _Charset; }
		}

		/// <summary>
		/// The location of 'Sub' or 'Function'.
		/// </summary>
		public Location SubOrFunctionLocation {
			get { return _SubOrFunctionLocation; }
		}

		/// <summary>
		/// The location of 'Lib', if any.
		/// </summary>
		public Location LibLocation {
			get { return _LibLocation; }
		}

		/// <summary>
		/// The library, if any.
		/// </summary>
		public StringLiteralExpression LibLiteral {
			get { return _LibLiteral; }
		}

		/// <summary>
		/// The location of 'Alias', if any.
		/// </summary>
		public Location AliasLocation {
			get { return _AliasLocation; }
		}

		/// <summary>
		/// The alias, if any.
		/// </summary>
		public StringLiteralExpression AliasLiteral {
			get { return _AliasLiteral; }
		}

		protected ExternalDeclaration(TreeType type, AttributeBlockCollection attributes, ModifierCollection modifiers, Location keywordLocation, Location charsetLocation, Charset charset, Location subOrFunctionLocation, SimpleName name, Location libLocation, StringLiteralExpression libLiteral, 
		Location aliasLocation, StringLiteralExpression aliasLiteral, ParameterCollection parameters, Location asLocation, AttributeBlockCollection resultTypeAttributes, TypeName resultType, Span span, IList<Comment> comments) : base(type, attributes, modifiers, keywordLocation, name, null, parameters, asLocation, resultTypeAttributes, resultType, 
		span, comments)
		{

			SetParent(libLiteral);
			SetParent(aliasLiteral);

			_CharsetLocation = charsetLocation;
			_Charset = charset;
			_SubOrFunctionLocation = subOrFunctionLocation;
			_LibLocation = libLocation;
			_LibLiteral = libLiteral;
			_AliasLocation = aliasLocation;
			_AliasLiteral = aliasLiteral;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			base.GetChildTrees(childList);
			AddChild(childList, LibLiteral);
			AddChild(childList, AliasLiteral);
		}
	}
}

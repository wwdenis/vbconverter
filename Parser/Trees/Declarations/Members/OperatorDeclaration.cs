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
using VBConverter.CodeParser.Tokens;
using VBConverter.CodeParser.Trees.Attributes;
using VBConverter.CodeParser.Trees.Comments;
using VBConverter.CodeParser.Trees.Declarations;
using VBConverter.CodeParser.Trees.Modifiers;
using VBConverter.CodeParser.Trees.Parameters;
using VBConverter.CodeParser.Trees.Statements;
using VBConverter.CodeParser.Trees.TypeNames;

namespace VBConverter.CodeParser.Trees.Declarations.Members
{
	/// <summary>
	/// A parse tree for an overloaded operator declaration.
	/// </summary>
	public sealed class OperatorDeclaration : MethodDeclaration
	{

		private readonly Token _OperatorToken;

		/// <summary>
		/// The operator being overloaded.
		/// </summary>
		public Token OperatorToken {
			get { return _OperatorToken; }
		}

		/// <summary>
		/// Creates a new parse tree for an overloaded operator declaration.
		/// </summary>
		/// <param name="attributes">The attributes for the parse tree.</param>
		/// <param name="modifiers">The modifiers for the parse tree.</param>
		/// <param name="keywordLocation">The location of the keyword.</param>
		/// <param name="operatorToken">The operator being overloaded.</param>
		/// <param name="parameters">The parameters of the declaration.</param>
		/// <param name="asLocation">The location of the 'As', if any.</param>
		/// <param name="resultTypeAttributes">The attributes on the result type, if any.</param>
		/// <param name="resultType">The result type, if any.</param>
		/// <param name="statements">The statements in the declaration.</param>
		/// <param name="endDeclaration">The end block declaration, if any.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public OperatorDeclaration(AttributeBlockCollection attributes, ModifierCollection modifiers, Location keywordLocation, Token operatorToken, ParameterCollection parameters, Location asLocation, AttributeBlockCollection resultTypeAttributes, TypeName resultType, StatementCollection statements, EndBlockDeclaration endDeclaration, 
		Span span, IList<Comment> comments) : base(TreeType.OperatorDeclaration, attributes, modifiers, keywordLocation, null, null, parameters, asLocation, resultTypeAttributes, resultType, 
		null, null, statements, endDeclaration, span, comments)
		{

			_OperatorToken = operatorToken;
		}
	}
}

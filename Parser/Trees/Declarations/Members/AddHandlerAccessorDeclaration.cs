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
using VBConverter.CodeParser.Trees.Parameters;
using VBConverter.CodeParser.Trees.Statements;

namespace VBConverter.CodeParser.Trees.Declarations.Members
{
	/// <summary>
	/// A parse tree for a AddHandler property accessor.
	/// </summary>
	public sealed class AddHandlerAccessorDeclaration : ModifiedDeclaration
	{

		private readonly Location _AddHandlerLocation;
		private readonly ParameterCollection _Parameters;
		private readonly StatementCollection _Statements;
		private readonly EndBlockDeclaration _EndStatement;

		/// <summary>
		/// The location of the 'AddHandler'.
		/// </summary>
		public Location AddHandlerLocation {
			get { return _AddHandlerLocation; }
		}

		/// <summary>
		/// The accessor's parameters.
		/// </summary>
		public ParameterCollection Parameters {
			get { return _Parameters; }
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
		public EndBlockDeclaration EndStatement {
			get { return _EndStatement; }
		}

		/// <summary>
		/// Constructs a new parse tree for a property accessor.
		/// </summary>
		/// <param name="attributes">The attributes for the parse tree.</param>
		/// <param name="addHandlerLocation">The location of the 'AddHandler'.</param>
		/// <param name="parameters">The parameters of the declaration.</param>
		/// <param name="statements">The statements in the declaration.</param>
		/// <param name="endStatement">The end block declaration, if any.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public AddHandlerAccessorDeclaration(AttributeBlockCollection attributes, Location addHandlerLocation, ParameterCollection parameters, StatementCollection statements, EndBlockDeclaration endStatement, Span span, IList<Comment> comments) : base(TreeType.AddHandlerAccessorDeclaration, attributes, null, span, comments)
		{

			SetParent(parameters);
			SetParent(statements);
			SetParent(endStatement);

			_Parameters = parameters;
			_AddHandlerLocation = addHandlerLocation;
			_Statements = statements;
			_EndStatement = endStatement;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			base.GetChildTrees(childList);

			AddChild(childList, Parameters);
			AddChild(childList, Statements);
			AddChild(childList, EndStatement);
		}
	}
}

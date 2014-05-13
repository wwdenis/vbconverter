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
using VBConverter.CodeParser.Trees.Parameters;
using VBConverter.CodeParser.Trees.Statements;

namespace VBConverter.CodeParser.Trees.Declarations.Members
{
	/// <summary>
	/// A parse tree for a Set property accessor.
	/// </summary>
	public sealed class SetAccessorDeclaration : ModifiedDeclaration
	{

		private readonly Location _SetLocation;
		private readonly ParameterCollection _Parameters;
		private readonly StatementCollection _Statements;
		private readonly EndBlockDeclaration _EndDeclaration;

		/// <summary>
		/// The location of the 'Set'.
		/// </summary>
		public Location SetLocation {
			get { return _SetLocation; }
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
		public EndBlockDeclaration EndDeclaration {
			get { return _EndDeclaration; }
		}

		/// <summary>
		/// Constructs a new parse tree for a property accessor.
		/// </summary>
		/// <param name="attributes">The attributes for the parse tree.</param>
		/// <param name="modifiers">The modifiers for the parse tree.</param>
		/// <param name="setLocation">The location of the 'Set'.</param>
		/// <param name="parameters">The parameters of the declaration.</param>
		/// <param name="statements">The statements in the declaration.</param>
		/// <param name="endDeclaration">The end block declaration, if any.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public SetAccessorDeclaration(AttributeBlockCollection attributes, ModifierCollection modifiers, Location setLocation, ParameterCollection parameters, StatementCollection statements, EndBlockDeclaration endDeclaration, Span span, IList<Comment> comments) : base(TreeType.SetAccessorDeclaration, attributes, modifiers, span, comments)
		{

			SetParent(parameters);
			SetParent(statements);
			SetParent(endDeclaration);

			_Parameters = parameters;
			_SetLocation = setLocation;
			_Statements = statements;
			_EndDeclaration = endDeclaration;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			base.GetChildTrees(childList);

			AddChild(childList, Parameters);
			AddChild(childList, Statements);
			AddChild(childList, EndDeclaration);
		}
	}
}

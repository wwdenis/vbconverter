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
using VBConverter.CodeParser.Trees.Arguments;
using VBConverter.CodeParser.Trees.TypeNames;

namespace VBConverter.CodeParser.Trees.Expressions
{
	/// <summary>
	/// A parse tree for a New expression.
	/// </summary>
	public sealed class NewExpression : Expression
	{

		private readonly TypeName _Target;
		private readonly ArgumentCollection _Arguments;

		/// <summary>
		/// The target type to create.
		/// </summary>
		public TypeName Target {
			get { return _Target; }
		}

		/// <summary>
		/// The arguments to the constructor.
		/// </summary>
		public ArgumentCollection Arguments {
			get { return _Arguments; }
		}

		/// <summary>
		/// Constructs a new parse tree for a New expression.
		/// </summary>
		/// <param name="target">The target type to create.</param>
		/// <param name="arguments">The arguments to the constructor.</param>
		/// <param name="span">The location of the parse tree.</param>
		public NewExpression(TypeName target, ArgumentCollection arguments, Span span) : base(TreeType.NewExpression, span)
		{

			if (target == null)
			{
				throw new ArgumentNullException("target");
			}

			SetParent(target);
			SetParent(arguments);

			_Target = target;
			_Arguments = arguments;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, Target);
			AddChild(childList, Arguments);
		}
	}
}

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
using VBConverter.CodeParser.Trees.Initializers;
using VBConverter.CodeParser.Trees.TypeNames;

namespace VBConverter.CodeParser.Trees.Expressions
{
	/// <summary>
	/// A parse tree for a New array expression.
	/// </summary>
	public sealed class NewAggregateExpression : Expression
	{

		private readonly ArrayTypeName _Target;
		private readonly AggregateInitializer _Initializer;

		/// <summary>
		/// The target array type to create.
		/// </summary>
		public ArrayTypeName Target {
			get { return _Target; }
		}

		/// <summary>
		/// The initializer for the array.
		/// </summary>
		public AggregateInitializer Initializer {
			get { return _Initializer; }
		}

		/// <summary>
		/// The constructor for a New array expression parse tree.
		/// </summary>
		/// <param name="target">The target array type to create.</param>
		/// <param name="initializer">The initializer for the array.</param>
		/// <param name="span">The location of the parse tree.</param>
		public NewAggregateExpression(ArrayTypeName target, AggregateInitializer initializer, Span span) : base(TreeType.NewAggregateExpression, span)
		{

			if (target == null)
			{
				throw new ArgumentNullException("target");
			}

			if (initializer == null)
			{
				throw new ArgumentNullException("initializer");
			}

			SetParent(target);
			SetParent(initializer);

			_Target = target;
			_Initializer = initializer;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, Target);
			AddChild(childList, Initializer);
		}
	}
}

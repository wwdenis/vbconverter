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
namespace VBConverter.CodeParser.Trees.Initializers
{
	/// <summary>
	/// A parse tree for an aggregate initializer.
	/// </summary>
	public sealed class AggregateInitializer : Initializer
	{

		private readonly InitializerCollection _Elements;

		/// <summary>
		/// The elements of the aggregate initializer.
		/// </summary>
		public InitializerCollection Elements {
			get { return _Elements; }
		}

		/// <summary>
		/// Constructs a new aggregate initializer parse tree.
		/// </summary>
		/// <param name="elements">The elements of the aggregate initializer.</param>
		/// <param name="span">The location of the parse tree.</param>
		public AggregateInitializer(InitializerCollection elements, Span span) : base(TreeType.AggregateInitializer, span)
		{

			if (elements == null)
			{
				throw new ArgumentNullException("elements");
			}

			SetParent(elements);
			_Elements = elements;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, Elements);
		}
	}
}

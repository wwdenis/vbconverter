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
using VBConverter.CodeParser.Trees.Collections;

namespace VBConverter.CodeParser.Trees.Attributes
{
	/// <summary>
	/// A read-only collection of attributes.
	/// </summary>
	public sealed class AttributeBlockCollection : TreeCollection<AttributeCollection>
	{

		/// <summary>
		/// Constructs a new collection of attribute blocks.
		/// </summary>
		/// <param name="attributeBlocks">The attribute blockss in the collection.</param>
		/// <param name="span">The location of the parse tree.</param>
		public AttributeBlockCollection(IList<AttributeCollection> attributeBlocks, Span span) : base(TreeType.AttributeBlockCollection, attributeBlocks, span)
		{

			if (attributeBlocks == null || attributeBlocks.Count == 0)
			{
				throw new ArgumentException("AttributeBlocksCollection cannot be empty.");
			}
		}
	}
}

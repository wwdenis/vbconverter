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
using VBConverter.CodeParser.Trees.Comments;
using VBConverter.CodeParser.Trees.Types;

namespace VBConverter.CodeParser.Trees.Statements
{
	/// <summary>
	/// A parse tree for an End statement of a block.
	/// </summary>
	public sealed class EndBlockStatement : Statement
	{

		private readonly BlockType _EndType;
		private readonly Location _EndArgumentLocation;

		/// <summary>
		/// The type of block the statement ends.
		/// </summary>
		public BlockType EndType {
			get { return _EndType; }
		}

		/// <summary>
		/// The location of the end block argument.
		/// </summary>
		public Location EndArgumentLocation {
			get { return _EndArgumentLocation; }
		}

		/// <summary>
		/// Creates a new parse tree for an End block statement.
		/// </summary>
		/// <param name="endType">The type of the block the statement ends.</param>
		/// <param name="endArgumentLocation">The location of the end block argument.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public EndBlockStatement(BlockType endType, Location endArgumentLocation, Span span, IList<Comment> comments) : base(TreeType.EndBlockStatement, span, comments)
		{

			if (endType < BlockType.While || endType > BlockType.Namespace)
				throw new ArgumentOutOfRangeException("endType");
			
			_EndType = endType;
			_EndArgumentLocation = endArgumentLocation;
		}
	}
}

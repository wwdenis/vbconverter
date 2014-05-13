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
	/// A parse tree for an Continue statement.
	/// </summary>
	public sealed class ContinueStatement : Statement
	{

		private readonly BlockType _ContinueType;
		private readonly Location _ContinueArgumentLocation;

		/// <summary>
		/// The type of tree this statement continues.
		/// </summary>
		public BlockType ContinueType {
			get { return _ContinueType; }
		}

		/// <summary>
		/// The location of the Continue statement type.
		/// </summary>
		public Location ContinueArgumentLocation {
			get { return _ContinueArgumentLocation; }
		}

		/// <summary>
		/// Constructs a parse tree for an Continue statement.
		/// </summary>
		/// <param name="continueType">The type of tree this statement continues.</param>
		/// <param name="continueArgumentLocation">The location of the Continue statement type.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public ContinueStatement(BlockType continueType, Location continueArgumentLocation, Span span, IList<Comment> comments) : base(TreeType.ContinueStatement, span, comments)
		{

			switch (continueType) {
				case BlockType.Do:
				case BlockType.For:
				case BlockType.While:
				case BlockType.None:
					break;
				// OK

				default:
					throw new ArgumentOutOfRangeException("continueType");
			}

			_ContinueType = continueType;
			_ContinueArgumentLocation = continueArgumentLocation;
		}
	}
}

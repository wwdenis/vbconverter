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
using VBConverter.CodeParser.Trees.Statements;
using VBConverter.CodeParser.Trees.Types;

namespace VBConverter.CodeParser.Trees.Declarations
{
	/// <summary>
	/// A parse tree for an End declaration.
	/// </summary>
	public sealed class EndBlockDeclaration : Declaration
	{

		private readonly BlockType _EndType;
		private readonly Location _EndArgumentLocation;

		/// <summary>
		/// The type of block the declaration ends.
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
		/// Creates a new parse tree for an End block declaration.
		/// </summary>
		/// <param name="endType">The type of the block the statement ends.</param>
		/// <param name="endArgumentLocation">The location of the end block argument.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public EndBlockDeclaration(BlockType endType, Location endArgumentLocation, Span span, IList<Comment> comments) : base(TreeType.EndBlockDeclaration, span, comments)
		{

			if (endType < BlockType.Sub && endType > BlockType.Namespace)
				throw new ArgumentOutOfRangeException("endType");
			
			_EndType = endType;
			_EndArgumentLocation = endArgumentLocation;
		}

		internal EndBlockDeclaration(EndBlockStatement endBlockStatement) : base(TreeType.EndBlockDeclaration, endBlockStatement.Span, endBlockStatement.Comments)
		{

			// We only need to convert these types.
			switch (endBlockStatement.EndType) {
				case BlockType.Function:
				case BlockType.Get:
				case BlockType.Set:
				case BlockType.Sub:
				case BlockType.Operator:
				case BlockType.AddHandler:
				case BlockType.RemoveHandler:
				case BlockType.RaiseEvent:
                case BlockType.Property:
					_EndType = endBlockStatement.EndType;
					break;

				default:
					throw new ArgumentException("Invalid EndBlockStatement type.");
			}

			_EndArgumentLocation = endBlockStatement.EndArgumentLocation;
		}
	}
}

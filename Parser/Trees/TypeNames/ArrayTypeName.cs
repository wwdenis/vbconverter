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

namespace VBConverter.CodeParser.Trees.TypeNames
{
	/// <summary>
	/// A parse tree for an array type name.
	/// </summary>
	/// <remarks>
	/// This tree may contain size arguments as well.
	/// </remarks>
	public sealed class ArrayTypeName : TypeName
	{

		private readonly TypeName _ElementTypeName;
		private readonly int _Rank;
		private readonly ArgumentCollection _Arguments;

		/// <summary>
		/// The type name for the element type of the array.
		/// </summary>
		public TypeName ElementTypeName {
			get { return _ElementTypeName; }
		}

		/// <summary>
		/// The rank of the array type name.
		/// </summary>
		public int Rank {
			get { return _Rank; }
		}

		/// <summary>
		/// The arguments of the array type name, if any.
		/// </summary>
		public ArgumentCollection Arguments {
			get { return _Arguments; }
		}

		/// <summary>
		/// Constructs a new parse tree for an array type name.
		/// </summary>
		/// <param name="elementTypeName">The type name for the array element type.</param>
		/// <param name="rank">The rank of the array type name.</param>
		/// <param name="arguments">The arguments of the array type name, if any.</param>
		/// <param name="span">The location of the parse tree.</param>
		public ArrayTypeName(TypeName elementTypeName, int rank, ArgumentCollection arguments, Span span) : base(TreeType.ArrayType, span)
		{

			if (arguments == null)
			{
				throw new ArgumentNullException("arguments");
			}

			SetParent(elementTypeName);
			SetParent(arguments);

			_ElementTypeName = elementTypeName;
			_Rank = rank;
			_Arguments = arguments;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, ElementTypeName);
			AddChild(childList, Arguments);
		}
	}
}

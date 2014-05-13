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

namespace VBConverter.CodeParser.Trees.Declarations
{
	/// <summary>
	/// A parse tree for an Option declaration.
	/// </summary>
	public sealed class OptionDeclaration : Declaration
	{

		private readonly OptionType _OptionType;
		private readonly Location _OptionTypeLocation;
		private readonly Location _OptionArgumentLocation;

		/// <summary>
		/// The type of Option statement.
		/// </summary>
		public OptionType OptionType {
			get { return _OptionType; }
		}

		/// <summary>
		/// The location of the Option type (e.g. "Strict"), if any.
		/// </summary>
		public Location OptionTypeLocation {
			get { return _OptionTypeLocation; }
		}

		/// <summary>
		/// The location of the Option argument (e.g. "On"), if any.
		/// </summary>
		public Location OptionArgumentLocation {
			get { return _OptionArgumentLocation; }
		}

		/// <summary>
		/// Constructs a new parse tree for an Option declaration.
		/// </summary>
		/// <param name="optionType">The type of the Option declaration.</param>
		/// <param name="optionTypeLocation">The location of the Option type, if any.</param>
		/// <param name="optionArgumentLocation">The location of the Option argument, if any.</param>
		/// <param name="span">The location of the parse tree.</param>
		/// <param name="comments">The comments for the parse tree.</param>
		public OptionDeclaration(OptionType optionType, Location optionTypeLocation, Location optionArgumentLocation, Span span, IList<Comment> comments) : base(TreeType.OptionDeclaration, span, comments)
		{

            if (!Enum.IsDefined(typeof(OptionType), optionType))
				throw new ArgumentOutOfRangeException("optionType");
			
			_OptionType = optionType;
			_OptionTypeLocation = optionTypeLocation;
			_OptionArgumentLocation = optionArgumentLocation;
		}
	}
}

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
namespace VBConverter.CodeParser.Trees.Modifiers
{
	/// <summary>
	/// A parse tree for a declaration modifier.
	/// </summary>
	public sealed class Modifier : Tree
	{

		private readonly ModifierTypes _ModifierType;

		/// <summary>
		/// The type of the modifier.
		/// </summary>
		public ModifierTypes ModifierType {
			get { return _ModifierType; }
		}

		/// <summary>
		/// Constructs a new modifier parse tree.
		/// </summary>
		/// <param name="modifierType">The type of the modifier.</param>
		/// <param name="span">The location of the parse tree.</param>
		public Modifier(ModifierTypes modifierType, Span span) : base(TreeType.Modifier, span)
		{

			if ((modifierType & (modifierType - 1)) != 0 || modifierType < ModifierTypes.None || modifierType > ModifierTypes.Narrowing)
				throw new ArgumentOutOfRangeException("modifierType");

			_ModifierType = modifierType;
		}
	}
}

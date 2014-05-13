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

namespace VBConverter.CodeParser.Trees.Modifiers
{
	/// <summary>
	/// A read-only collection of modifiers.
	/// </summary>
	public sealed class ModifierCollection : TreeCollection<Modifier>
	{

		private readonly ModifierTypes _ModifierTypes;

		/// <summary>
		/// All the modifiers in the collection.
		/// </summary>
		public ModifierTypes ModifierTypes {
			get { return _ModifierTypes; }
		}

		/// <summary>
		/// Constructs a collection of modifiers.
		/// </summary>
		/// <param name="modifiers">The modifiers in the collection.</param>
		/// <param name="span">The location of the parse tree.</param>
		public ModifierCollection(IList<Modifier> modifiers, Span span) : base(TreeType.ModifierCollection, modifiers, span)
		{

			if (modifiers == null || modifiers.Count == 0)
				throw new ArgumentException("ModifierCollection cannot be empty.");

			foreach (Modifier Modifier in modifiers)
				_ModifierTypes = _ModifierTypes | Modifier.ModifierType;
		}

        public bool IsOfType(ModifierTypes type)
        {
            return ((_ModifierTypes & type) == type);
        }
	}
}

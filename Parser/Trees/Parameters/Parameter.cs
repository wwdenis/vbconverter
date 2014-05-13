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
using VBConverter.CodeParser.Trees.Attributes;
using VBConverter.CodeParser.Trees.Initializers;
using VBConverter.CodeParser.Trees.Modifiers;
using VBConverter.CodeParser.Trees.Names;
using VBConverter.CodeParser.Trees.TypeNames;

namespace VBConverter.CodeParser.Trees.Parameters
{
	/// <summary>
	/// A parse tree for a parameter.
	/// </summary>
	public sealed class Parameter : Tree
	{

		private readonly AttributeBlockCollection _Attributes;
		private readonly ModifierCollection _Modifiers;
		private readonly VariableName _VariableName;
		private readonly Location _AsLocation;
		private readonly TypeName _ParameterType;
		private readonly Location _EqualsLocation;
		private readonly Initializer _Initializer;

		/// <summary>
		/// The attributes on the parameter.
		/// </summary>
		public AttributeBlockCollection Attributes {
			get { return _Attributes; }
		}

		/// <summary>
		/// The modifiers on the parameter.
		/// </summary>
		public ModifierCollection Modifiers {
			get { return _Modifiers; }
		}

		/// <summary>
		/// The name of the parameter.
		/// </summary>
		public VariableName VariableName {
			get { return _VariableName; }
		}

		/// <summary>
		/// The location of the 'As', if any.
		/// </summary>
		public Location AsLocation {
			get { return _AsLocation; }
		}

		/// <summary>
		/// The parameter type, if any.
		/// </summary>
		public TypeName ParameterType {
			get { return _ParameterType; }
		}

		/// <summary>
		/// The location of the '=', if any.
		/// </summary>
		public Location EqualsLocation {
			get { return _EqualsLocation; }
		}

		/// <summary>
		/// The initializer for the parameter, if any.
		/// </summary>
		public Initializer Initializer {
			get { return _Initializer; }
		}

		/// <summary>
		/// Constructs a new parameter parse tree.
		/// </summary>
		/// <param name="attributes">The attributes on the parameter.</param>
		/// <param name="modifiers">The modifiers on the parameter.</param>
		/// <param name="variableName">The name of the parameter.</param>
		/// <param name="asLocation">The location of the 'As'.</param>
		/// <param name="parameterType">The type of the parameter. Can be Nothing.</param>
		/// <param name="equalsLocation">The location of the '='.</param>
		/// <param name="initializer">The initializer for the parameter. Can be Nothing.</param>
		/// <param name="span">The location of the parse tree.</param>
		public Parameter(AttributeBlockCollection attributes, ModifierCollection modifiers, VariableName variableName, Location asLocation, TypeName parameterType, Location equalsLocation, Initializer initializer, Span span) : base(TreeType.Parameter, span)
		{

			if (variableName == null)
			{
				throw new ArgumentNullException("variableName");
			}

			SetParent(attributes);
			SetParent(modifiers);
			SetParent(variableName);
			SetParent(parameterType);
			SetParent(initializer);

			_Attributes = attributes;
			_Modifiers = modifiers;
			_VariableName = variableName;
			_AsLocation = asLocation;
			_ParameterType = parameterType;
			_EqualsLocation = equalsLocation;
			_Initializer = initializer;
		}

		protected override void GetChildTrees(IList<Tree> childList)
		{
			AddChild(childList, Attributes);
			AddChild(childList, Modifiers);
			AddChild(childList, VariableName);
			AddChild(childList, ParameterType);
			AddChild(childList, Initializer);
		}
	}
}

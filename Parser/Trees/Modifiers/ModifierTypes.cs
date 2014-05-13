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
	/// The type of a parse tree modifier.
	/// </summary>
	[Flags()]
	public enum ModifierTypes
	{
		None = 0,
		Public = 1,
		Private = 2,
		Protected = 4,
		Friend = 8,
		AccessModifiers = Public | Private | Protected | Friend,
		Static = 16,
		Shared = 32,
		Shadows = 64,
		Overloads = 128,
		MustInherit = 256,
		NotInheritable = 512,
		Overrides = 1024,
		NotOverridable = 2048,
		Overridable = 4096,
		MustOverride = 8192,
		ReadOnly = 16384,
		WriteOnly = 32768,
		Dim = 65536,
		Const = 131072,
		Default = 262144,
		WithEvents = 524288,
		ByVal = 1048576,
		ByRef = 2097152,
		Optional = 4194304,
		ParamArray = 8388608,
		Partial = 16777216,
		Widening = 33554432,
		Narrowing = 67108864
	}
}

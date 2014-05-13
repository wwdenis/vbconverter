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
namespace VBConverter.CodeParser.Trees.Types
{
	/// <summary>
	/// The type of an intrinsic type name.
	/// </summary>
	public enum IntrinsicType
	{
		None,
        Boolean,
		SByte,
		Byte,
		Short,
		UShort,
		Integer,
		UInteger,
		Long,
		ULong,
		Decimal,
		Single,
		Double,
		Date,
		Char,
		String,
		Object, 
        Variant,
        Currency,
        FixedString
	}
}

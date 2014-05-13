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
namespace VBConverter.CodeParser.Tokens
{
	/// <summary>
	/// The type of a token.
	/// </summary>
	public enum TokenType
	{
        None,
		LexicalError,
		EndOfStream,
		LineTerminator,
		Comment,
        // Identifier
        Identifier,
        // Literals
		StringLiteral,
		CharacterLiteral,
		DateLiteral,
		IntegerLiteral,
		UnsignedIntegerLiteral,
		FloatingPointLiteral,
		DecimalLiteral,
        // Keywords
		AddHandler,
		AddressOf,
		Alias,
		And,
		AndAlso,
		Ansi,
		As,
		Assembly,
		Auto,
        Base,
		Binary,
		Boolean,
		ByRef,
		Byte,
		ByVal,
		Call,
		Case,
		Catch,
		CBool,
		CByte,
		CChar,
        CCur,
		CDate,
		CDec,
		CDbl,
		Char,
		CInt,
		Class,
		CLng,
		CObj,
		Compare,
		Const,
		Continue,
		CSByte,
		CShort,
		CSng,
		CStr,
        CType,
		CUInt,
		CULng,
        Currency,
		CUShort,
		Custom,
        CVar,
		Date,
		Decimal,
		Declare,
		Default,
		Delegate,
		Dim,
		DirectCast,
		Do,
		Double,
		Each,
		Else,
		ElseIf,
		End,
		EndIf,
		Enum,
		Erase,
		Error,
		Event,
		Exit,
		Explicit,
		ExternalChecksum,
		ExternalSource,
		False,
		Finally,
		For,
		Friend,
		Function,
		Get,
		GetType,
		Global,
		GoSub,
		GoTo,
		Handles,
		If,
		Implements,
		Imports,
		In,
		Inherits,
		Integer,
		Interface,
		Is,
		IsTrue,
		IsNot,
		IsFalse,
		Let,
		Lib,
		Like,
		Long,
		Loop,
		Me,
		Mid,
		Mod,
		Module,
		MustInherit,
		MustOverride,
		MyBase,
		MyClass,
		Namespace,
		Narrowing,
		New,
		Next,
		Not,
		Nothing,
		NotInheritable,
		NotOverridable,
		Object,
		Of,
		Off,
		On,
		Operator,
		Option,
		Optional,
		Or,
		OrElse,
		Overloads,
		Overridable,
		Overrides,
		ParamArray,
		Partial,
		Preserve,
		Private,
		Property,
		Protected,
		Public,
		RaiseEvent,
		ReadOnly,
		ReDim,
		Region,
		Rem,
		RemoveHandler,
		Resume,
		Return,
		SByte,
		Select,
		Set,
		Shadows,
		Shared,
		Short,
		Single,
		Static,
		Step,
		Stop,
		Strict,
		String,
		Structure,
		Sub,
		SyncLock,
		Text,
		Then,
		Throw,
		To,
		True,
		Try,
		TryCast,
        Type,
		TypeOf,
		UInteger,
		ULong,
		UShort,
		Using,
		Unicode,
		Until,
		Variant,
		Wend,
		When,
		While,
		Widening,
		With,
		WithEvents,
		WriteOnly,
		Xor,
        //Punctuators
		LeftParenthesis,    // (
		RightParenthesis,   // )
		LeftCurlyBrace,     // {
		RightCurlyBrace,    // }
		Exclamation,        // !
		Pound,              // #
		Comma,              // ,
		Period,             // .
        Colon,              // :
		ColonEquals,        // :=
		Ampersand,          // &
		AmpersandEquals,    // &=
		Star,               // *
		StarEquals,         // *=
		Plus,               // +
		PlusEquals,         // +=
		Minus,              // -
		MinusEquals,        // -=
		ForwardSlash,       // /
		ForwardSlashEquals, // /=
		BackwardSlash,      // \
		BackwardSlashEquals,// \=
		Caret,              // ^
		CaretEquals,        // ^=
		LessThan,           // <
		LessThanEquals,     // <=
		Equals,             // =
		NotEquals,          // <>
		GreaterThan,        // >
		GreaterThanEquals,  // >=
		LessThanLessThan,   // <<
		LessThanLessThanEquals,         // <<=
		GreaterThanGreaterThan,         // >>
		GreaterThanGreaterThanEquals    // >>=
	}
}

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

namespace VBConverter.CodeParser.Base
{
    /// <summary>
    /// A character that denotes the type of something.
    /// </summary>
    [Flags()]
    public enum TypeCharacter
    {
        /// <summary>No type character</summary>
        None = 0,

        /// <summary>The String symbol '$'.</summary>
        StringSymbol = 1,

        /// <summary>The Integer symbol '%'.</summary>
        IntegerSymbol = 2,

        /// <summary>The Long symbol '&amp;'.</summary>
        LongSymbol = 4,

        /// <summary>The Short character 'S'.</summary>
        ShortChar = 8,

        /// <summary>The Integer character 'I'.</summary>
        IntegerChar = 16,

        /// <summary>The Long character 'L'.</summary>
        LongChar = 32,

        /// <summary>The Single symbol '!'.</summary>
        SingleSymbol = 64,

        /// <summary>The Double symbol '#'.</summary>
        DoubleSymbol = 128,

        /// <summary>The Decimal symbol '@'.</summary>
        DecimalSymbol = 256,

        /// <summary>The Single character 'F'.</summary>
        SingleChar = 512,

        /// <summary>The Double character 'R'.</summary>
        DoubleChar = 1024,

        /// <summary>The Decimal character 'D'.</summary>
        DecimalChar = 2048,

        /// <summary>The unsigned Short characters 'US'.</summary>
        /// <remarks>New for Visual Basic 8.0.</remarks>
        UnsignedShortChar = 4096,

        /// <summary>The unsigned Integer characters 'UI'.</summary>
        /// <remarks>New for Visual Basic 8.0.</remarks>
        UnsignedIntegerChar = 8192,

        /// <summary>The unsigned Long characters 'UL'.</summary>
        /// <remarks>New for Visual Basic 8.0.</remarks>
        UnsignedLongChar = 16384,

        /// <summary>All type characters.</summary>
        All = 32767
    }
}
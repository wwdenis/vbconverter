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
    /// The numeric base of an integer literal.
    /// </summary>
    public enum IntegerBase
    {
        /// <summary>Base 10.</summary>
        Decimal,

        /// <summary>Base 8.</summary>
        Octal,

        /// <summary>Base 16.</summary>
        Hexadecimal
    }
}
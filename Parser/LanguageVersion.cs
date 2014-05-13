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

namespace VBConverter.CodeParser
{
    /// <summary>
    /// The version of the language you want.
    /// </summary>
    [Flags()]
    public enum LanguageVersion
    {
        Non = 0,

        /// <summary>Visual Basic 6.0</summary>
        /// <remarks>Shipped in Visual Basic 6.0</remarks>
        VB6 = 1,

        /// <summary>Visual Basic 7.1</summary>
        /// <remarks>Shipped in Visual Basic 2003</remarks>
        VB7 = 2,

        /// <summary>Visual Basic 8.0</summary>
        /// <remarks>Shipped in Visual Basic 2005</remarks>
        VB8 = 4,

        /// <summary>.NET Versions</summary>
        /// <remarks>All .NET versions</remarks>
        Net = 6,

        /// <summary>All Versions</summary>
        /// <remarks>All language versions</remarks>
        All = 7
    }
}
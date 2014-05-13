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

using VBConverter.CodeParser.Base;

namespace VBConverter.CodeParser.Trees.Names
{
    /// <summary>
    /// A parse tree for a simple name (e.g. 'foo').
    /// </summary>
    public sealed class SimpleName : Name
    {

        private readonly string _Name;
        private readonly TypeCharacter _TypeCharacter;
        private readonly bool _Escaped;

        /// <summary>
        /// The name, if any.
        /// </summary>
        public string Name
        {
            get { return _Name; }
        }

        /// <summary>
        /// The type character.
        /// </summary>
        public TypeCharacter TypeCharacter
        {
            get { return _TypeCharacter; }
        }

        /// <summary>
        /// Whether the name is escaped.
        /// </summary>
        public bool Escaped
        {
            get { return _Escaped; }
        }

        public override string FullName
        {
            get { return Name; }
        }

        /// <summary>
        /// Creates a bad simple name.
        /// </summary>
        /// <param name="Span">The location of the parse tree.</param>
        /// <returns>A bad simple name.</returns>
        public static SimpleName GetBadSimpleName(Span span)
        {
            return new SimpleName(span);
        }

        public override bool Equals(object obj)
        {
            if (obj != null && obj is SimpleName && ((SimpleName)obj).Name != null)
                return this.Name.Equals(((SimpleName)obj).Name, StringComparison.InvariantCultureIgnoreCase);
            else 
                return false;
        }

        public static bool operator ==(SimpleName first, SimpleName second)
        {
            return first.Equals(second);
        }

        public static bool operator !=(SimpleName first, SimpleName second)
        {
            return !first.Equals(second);
        }

        /// <summary>
        /// Constructs a new simple name parse tree.
        /// </summary>
        /// <param name="name">The name, if any.</param>
        /// <param name="typeCharacter">The type character.</param>
        /// <param name="escaped">Whether the name is escaped.</param>
        /// <param name="span">The location of the parse tree.</param>
        public SimpleName(string name, TypeCharacter typeCharacter, bool escaped, Span span) : base(TreeType.SimpleName, span)
        {

            if (typeCharacter != TypeCharacter.None && escaped)
                throw new ArgumentException("Escaped named cannot have type characters.");

            if (typeCharacter != TypeCharacter.None && typeCharacter != TypeCharacter.DecimalSymbol && typeCharacter != TypeCharacter.DoubleSymbol && typeCharacter != TypeCharacter.IntegerSymbol && typeCharacter != TypeCharacter.LongSymbol && typeCharacter != TypeCharacter.SingleSymbol && typeCharacter != TypeCharacter.StringSymbol)
                throw new ArgumentOutOfRangeException("typeCharacter");

            if (name == null)
                throw new ArgumentNullException("name");

            _Name = name;
            _TypeCharacter = typeCharacter;
            _Escaped = escaped;
        }

        private SimpleName(Span span) : base(TreeType.SimpleName, span)
        {
        }

        public override bool IsBad
        {
            get { return Name == null; }
        }
    }
}

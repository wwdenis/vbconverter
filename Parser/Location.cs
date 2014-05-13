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
    /// Stores source code line and column information.
    /// </summary>
    public struct Location
    {
        public static readonly Location Empty = new Location(0, 0, 0);

        private readonly long _Index;
        private readonly long _Line;
        private readonly long _Column;

        /// <summary>
        /// The index in the stream (0-based).
        /// </summary>
        public long Index
        {
            get { return _Index; }
        }

        /// <summary>
        /// The physical line number (1-based).
        /// </summary>
        public long Line
        {
            get { return _Line; }
        }

        /// <summary>
        /// The physical column number (1-based).
        /// </summary>
        public long Column
        {
            get { return _Column; }
        }

        /// <summary>
        /// Whether the location is a valid location.
        /// </summary>
        public bool IsValid
        {
            get { return Line != 0 && Column != 0; }
        }

        /// <summary>
        /// Constructs a new Location for a particular source location.
        /// </summary>
        /// <param name="index">The index in the stream (0-based).</param>
        /// <param name="line">The physical line number (1-based).</param>
        /// <param name="column">The physical column number (1-based).</param>
        public Location(long index, long line, long column)
        {
            _Index = index;
            _Line = line;
            _Column = column;
        }

        /// <summary>
        /// Compares two specified Location values to see if they are equal.
        /// </summary>
        /// <param name="left">One location to compare.</param>
        /// <param name="right">The other location to compare.</param>
        /// <returns>True if the locations are the same, False otherwise.</returns>
        public static bool operator ==(Location left, Location right)
        {
            //return left.Index == right.Index;
            return left.Line == right.Line && left.Column == right.Column;
        }

        /// <summary>
        /// Compares two specified Location values to see if they are not equal.
        /// </summary>
        /// <param name="left">One location to compare.</param>
        /// <param name="right">The other location to compare.</param>
        /// <returns>True if the locations are not the same, False otherwise.</returns>
        public static bool operator !=(Location left, Location right)
        {
            //return left.Index != right.Index;
            return left.Line != right.Line || left.Column != right.Column;
        }

        /// <summary>
        /// Compares two specified Location values to see if one is before the other.
        /// </summary>
        /// <param name="left">One location to compare.</param>
        /// <param name="right">The other location to compare.</param>
        /// <returns>True if the first location is before the other location, False otherwise.</returns>
        public static bool operator <(Location left, Location right)
        {
            return left.Index < right.Index;
        }

        /// <summary>
        /// Compares two specified Location values to see if one is after the other.
        /// </summary>
        /// <param name="left">One location to compare.</param>
        /// <param name="right">The other location to compare.</param>
        /// <returns>True if the first location is after the other location, False otherwise.</returns>
        public static bool operator >(Location left, Location right)
        {
            return left.Index > right.Index;
        }

        /// <summary>
        /// Compares two specified Location values to see if one is before or the same as the other.
        /// </summary>
        /// <param name="left">One location to compare.</param>
        /// <param name="right">The other location to compare.</param>
        /// <returns>True if the first location is before or the same as the other location, False otherwise.</returns>
        public static bool operator <=(Location left, Location right)
        {
            return left.Index <= right.Index;
        }

        /// <summary>
        /// Compares two specified Location values to see if one is after or the same as the other.
        /// </summary>
        /// <param name="left">One location to compare.</param>
        /// <param name="right">The other location to compare.</param>
        /// <returns>True if the first location is after or the same as the other location, False otherwise.</returns>
        public static bool operator >=(Location left, Location right)
        {
            return left.Index >= right.Index;
        }

        /// <summary>
        /// Compares two specified Location values.
        /// </summary>
        /// <param name="left">One location to compare.</param>
        /// <param name="right">The other location to compare.</param>
        /// <returns>0 if the locations are equal, -1 if the left one is less than the right one, 1 otherwise.</returns>
        public static int Compare(Location left, Location right)
        {
            if (left == right)
            {
                return 0;
            }
            else if (left < right)
            {
                return -1;
            }
            else
            {
                return 1;
            }
        }

        public override string ToString()
        {
            return "(" + this.Column + "," + this.Line + ")";
        }

        public override bool Equals(object obj)
        {
            if (obj is Location)
                return this == (Location)obj;
            else
                return false;
 
        }

        public override int GetHashCode()
        {
            // Mask off the upper 32 bits of the index and use that as
            // the hash code.
            return (int)Index & unchecked((int)4294967295L);
        }
    }
}
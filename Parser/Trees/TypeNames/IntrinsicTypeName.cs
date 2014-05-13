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

using VBConverter.CodeParser.Trees.Types;
using VBConverter.CodeParser.Trees.Expressions;

namespace VBConverter.CodeParser.Trees.TypeNames
{
    /// <summary>
    /// A parse tree for an intrinsic type name.
    /// </summary>
    public sealed class IntrinsicTypeName : TypeName
    {
        #region Fields

        private IntrinsicType _IntrinsicType;
        private Expression _stringLength = null;

        #endregion

        #region Properties

        /// <summary>
        /// The intrinsic type.
        /// </summary>
        public IntrinsicType IntrinsicType
        {
            get { return _IntrinsicType; }
        }

        /// <summary>
        /// The argument for fixed-length string (VB6)
        /// </summary>
        public Expression StringLength
        {
            get 
            { 
                return _stringLength; 
            }
            set
            {
                if (IntrinsicType != IntrinsicType.FixedString)
                    throw new InvalidOperationException("Cannot set length arguments for non-string types.");

                _stringLength = value;
            }
        }

        #endregion

        /// <summary>
        /// Constructs a new intrinsic type parse tree.
        /// </summary>
        /// <param name="intrinsicType">The intrinsic type.</param>
        /// <param name="span">The location of the parse tree.</param>
        public IntrinsicTypeName(IntrinsicType intrinsicType, Span span) : base(TreeType.IntrinsicType, span)
        {
            if (!Enum.IsDefined(typeof(IntrinsicType), intrinsicType))
                throw new ArgumentOutOfRangeException("intrinsicType");

            _IntrinsicType = intrinsicType;
        }
    }
}

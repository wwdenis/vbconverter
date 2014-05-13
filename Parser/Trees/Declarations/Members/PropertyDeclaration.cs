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
using VBConverter.CodeParser.Tokens;
using VBConverter.CodeParser.Trees.Attributes;
using VBConverter.CodeParser.Trees.Comments;
using VBConverter.CodeParser.Trees.Declarations;
using VBConverter.CodeParser.Trees.Modifiers;
using VBConverter.CodeParser.Trees.Names;
using VBConverter.CodeParser.Trees.Parameters;
using VBConverter.CodeParser.Trees.Statements;
using VBConverter.CodeParser.Trees.TypeNames;

namespace VBConverter.CodeParser.Trees.Declarations.Members
{
    /// <summary>
    /// A parse tree for a property declaration.
    /// </summary>
    public sealed class PropertyDeclaration : MethodDeclaration
    {
        private GetAccessorDeclaration _getAccessor;
        private SetAccessorDeclaration _setAccessor;

        /// <summary>
        /// The property get accessor
        /// </summary>
        public GetAccessorDeclaration GetAccessor
        {
            get 
            { 
                return _getAccessor; 
            }
            internal set
            {
                SetParent(value);
                _getAccessor = value;
            }
        }
        /// <summary>
        /// The property set accessor
        /// </summary>
        public SetAccessorDeclaration SetAccessor
        {
            get 
            { 
                return _setAccessor; 
            }
            internal set
            {
                SetParent(value);
                _setAccessor = value;
            }
        }

        /// <summary>
        /// Constructs a new parse tree for a property declaration (VB6).
        /// </summary>
        /// <param name="attributes">The attributes on the declaration.</param>
        /// <param name="modifiers">The modifiers on the declaration.</param>
        /// <param name="keywordLocation">The location of the keyword.</param>
        /// <param name="name">The name of the property.</param>
        /// <param name="parameters">The parameters of the property.</param>
        /// <param name="asLocation">The location of the 'As', if any.</param>
        /// <param name="resultTypeAttributes">The attributes on the result type.</param>
        /// <param name="resultType">The result type, if any.</param>
        /// <param name="implementsList">The implements list.</param>
        /// <param name="accessors">The property accessors.</param>
        /// <param name="statements">The property statements (VB6)</param>
        /// <param name="endDeclaration">The End Property declaration, if any.</param>
        /// <param name="span">The location of the parse tree.</param>
        /// <param name="comments">The comments for the parse tree.</param>
        public PropertyDeclaration(
            AttributeBlockCollection attributes,
            ModifierCollection modifiers,
            Location keywordLocation,
            SimpleName name,
            ParameterCollection parameters,
            Location asLocation,
            AttributeBlockCollection resultTypeAttributes,
            TypeName resultType,
            NameCollection implementsList,
            GetAccessorDeclaration getAccessor,
            SetAccessorDeclaration setAccessor,
            EndBlockDeclaration endDeclaration,
            Span span,
            IList<Comment> comments
        ) : 
        base(
            TreeType.PropertyDeclaration,
            attributes,
            modifiers,
            keywordLocation,
            name,
            null,
            parameters,
            asLocation,
            resultTypeAttributes,
            resultType,
            implementsList,
            null,
            null,
            endDeclaration,
            span,
            comments)
        {
            GetAccessor = getAccessor;
            SetAccessor = setAccessor;
        }

        protected override void GetChildTrees(IList<Tree> childList)
        {
            base.GetChildTrees(childList);
            AddChild(childList, GetAccessor);
            AddChild(childList, SetAccessor);
        }
    }
}

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
using VBConverter.CodeParser.Trees.Types;

namespace VBConverter.CodeParser.Tokens
{
    /// <summary>
    /// An identifier.
    /// </summary>
    public sealed class IdentifierToken : Token
    {
        #region Inner Elements

        private struct Keyword
        {
            public readonly LanguageVersion Versions;
            public readonly LanguageVersion ReservedVersions;
            public readonly TokenType TokenType;

            public Keyword(LanguageVersion Versions, LanguageVersion ReservedVersions, TokenType TokenType)
            {
                this.Versions = Versions;
                this.ReservedVersions = ReservedVersions;
                this.TokenType = TokenType;
            }
        }

        #endregion

        #region Fields

        private static Dictionary<string, Keyword> KeywordTable;
        private readonly string _Identifier;
        private readonly TokenType _UnreservedType;
        // Whether the identifier was escaped (i.e. [a])
        private readonly bool _Escaped;
        // The type character that followed, if any
        private readonly TypeCharacter _TypeCharacter;
        private readonly bool _hasSpaceAfter;

        #endregion

        #region Constructors

        static IdentifierToken()
        {
            BuildTable();
        }

        /// <summary>
        /// Constructs a new identifier token.
        /// </summary>
        /// <param name="type">The token type of the identifier.</param>
        /// <param name="unreservedType">The unreserved token type of the identifier.</param>
        /// <param name="identifier">The text of the identifier</param>
        /// <param name="escaped">Whether the identifier is escaped.</param>
        /// <param name="typeCharacter">The type character of the identifier.</param>
        /// <param name="span">The location of the identifier.</param>
        public IdentifierToken(TokenType type, TokenType unreservedType, string identifier, bool escaped, bool hasSpaceAfter, TypeCharacter typeCharacter, Span span) : base(type, span)
        {

            if (type != TokenType.Identifier && !IsKeyword(type))
                throw new ArgumentOutOfRangeException("type");

            if (unreservedType != TokenType.Identifier && !IsKeyword(unreservedType))
                throw new ArgumentOutOfRangeException("unreservedType");

            if (identifier == null || identifier == "")
                throw new ArgumentException("Identifier cannot be empty.", "identifier");

            if (typeCharacter != TypeCharacter.None && typeCharacter != TypeCharacter.DecimalSymbol && typeCharacter != TypeCharacter.DoubleSymbol && typeCharacter != TypeCharacter.IntegerSymbol && typeCharacter != TypeCharacter.LongSymbol && typeCharacter != TypeCharacter.SingleSymbol && typeCharacter != TypeCharacter.StringSymbol)
                throw new ArgumentOutOfRangeException("typeCharacter");

            if (typeCharacter != TypeCharacter.None && escaped)
                throw new ArgumentException("Escaped identifiers cannot have type characters.");

            _UnreservedType = unreservedType;
            _Identifier = identifier;
            _Escaped = escaped;
            _hasSpaceAfter = hasSpaceAfter;
            _TypeCharacter = typeCharacter;
        }
        #endregion

        #region Static Functions

        /// <summary>
        /// Returns the token type of the string.
        /// </summary>
        /// <param name="s">A string containing the token to convert</param>
        /// <param name="version">The version of Visual Basic. The version determines if a token is a keyword or not.</param>
        /// <param name="includeUnreserved">Defifes if a non-keyword token will be treated as a keyword.</param>
        /// <returns>A TokenType equivalent to the value contained in s.</returns>
        static public TokenType Parse(string s, LanguageVersion version, bool includeUnreserved)
        {
            if (KeywordTable.ContainsKey(s))
            {
                Keyword Table = KeywordTable[s];
                bool isKeyword = LanguageInfo.Implements(Table.Versions, version);
                bool isReserved = LanguageInfo.Implements(Table.ReservedVersions, version);

                if (isKeyword && (isReserved || includeUnreserved))
                    return Table.TokenType;
            }

            return TokenType.Identifier;
        }

        /// <summary>
        /// Determines if a token type is a keyword.
        /// </summary>
        /// <param name="type">The token type to check.</param>
        /// <returns>True if the token type is a keyword, False otherwise.</returns>
        public static bool IsKeyword(TokenType type)
        {
            return type >= TokenType.AddHandler && type <= TokenType.Xor;
        }

        /// <summary>
        /// Determines if a token type is a punctuator.
        /// </summary>
        /// <param name="type">The token type to check.</param>
        /// <returns>True if the token type is a punctuator, False otherwise.</returns>
        public static bool IsPunctuator(TokenType type)
        {
            return type >= TokenType.LeftParenthesis && type <= TokenType.GreaterThanGreaterThanEquals;
        }

        /// <summary>
        /// Determines if a token type is a literal.
        /// </summary>
        /// <param name="type">The token type to check.</param>
        /// <returns>True if the token type is a literal, False otherwise.</returns>
        public static bool IsLiteral(TokenType type)
        {
            return
                type >= TokenType.StringLiteral && type <= TokenType.DecimalLiteral ||
                type == TokenType.True ||
                type == TokenType.False;
        }

        /// <summary>
        /// Determines if a token type is a terminator.
        /// </summary>
        /// <param name="type">The token type to check.</param>
        /// <returns>True if the token type is a terminator, False otherwise.</returns>
        public static bool IsTerminator(TokenType type)
        {
            // return type >= TokenType.None && type <= TokenType.Comment;
            return
                type == TokenType.EndOfStream ||
                type == TokenType.Colon ||
                type == TokenType.LineTerminator ||
                type == TokenType.Comment;
        }

        /// <summary>
        /// Determines if a token type is a identifier.
        /// </summary>
        /// <param name="type">The token type to check.</param>
        /// <returns>True if the token type is a identifier, False otherwise.</returns>
        public static bool IsIdentifier(TokenType type)
        {
            return type == TokenType.Identifier && !IsTerminator(type);
        }

        /// <summary>
        /// Determines if a token type is a modifier.
        /// </summary>
        /// <param name="type">The token type to check.</param>
        /// <returns>True if the token type is a modifier, False otherwise.</returns>
        public static bool IsModifier(TokenType type)
        {
            return
                type == TokenType.Not ||
                type == TokenType.Minus;
        }

        /// <summary>
        /// Determines if a token type is a function.
        /// </summary>
        /// <param name="type">The token type to check.</param>
        /// <returns>True if the token type is a function, False otherwise.</returns>
        public static bool IsFunction(TokenType type)
        {
            //TODO: Verify all .NET instrisic functions
            //TODO: Verify Mid Statement (.NET) and Mid Function(VB6)
            TokenType[] ValidTokens = { 
                TokenType.CBool, 
                TokenType.CByte, 
                TokenType.CDate, 
                TokenType.CDbl, 
                TokenType.CDec, 
                TokenType.CInt, 
                TokenType.CLng, 
                TokenType.CObj, 
                TokenType.CSByte, 
                TokenType.CShort, 
                TokenType.CSng, 
                TokenType.CStr, 
                TokenType.CType, 
                TokenType.CUInt, 
                TokenType.CULng, 
                TokenType.CUShort, 
                TokenType.CVar, 
                TokenType.Mid,
                TokenType.DirectCast
            };

            bool Result = Array.Exists<TokenType>(ValidTokens, delegate(TokenType i) { return i == type; });

            return Result;
        }

        /// <summary>
        /// Determines if a token type is a special object, like Me, MyBase, etc..
        /// </summary>
        /// <param name="type">The token type to check.</param>
        /// <returns>True if the token type is a special object, False otherwise.</returns>
        public static bool IsSpecialObject(TokenType type)
        {
            //TODO: Verify all .NET instrisic SpecialObjects
            TokenType[] ValidTokens = { 
                TokenType.Me, 
                TokenType.MyBase, 
                TokenType.MyClass
            };

            bool Result = Array.Exists<TokenType>(ValidTokens, delegate(TokenType i) { return i == type; });

            return Result;
        }


        /// <summary>
        /// Determines if a token type can be evaluated (as a function or as a value).
        /// </summary>
        /// <param name="type">The token type to check.</param>
        /// <returns>True if the token type can be evaluated, False otherwise.</returns>
        public static bool IsEvaluable(TokenType type)
        {
            return
                IsIdentifier(type) ||
                IsLiteral(type) ||
                IsModifier(type) ||
                IsFunction(type) ||
                IsSpecialObject(type);
        }

        /// <summary>
        /// Determines if a token type is an assignemt accessor (Visual Basic 6).
        /// </summary>
        /// <param name="type">The token type to check.</param>
        /// <returns>True if the token type is a assignent accessor, False otherwise.</returns>
        public static bool IsAssignmentAccessor(TokenType type)
        {
            return
                type == TokenType.Let ||
                type == TokenType.Set;
        }

        public static IntrinsicType GetIntrinsicType(TokenType type)
        {
            IntrinsicType Result = IntrinsicType.None;
            switch (type)
            {
                case TokenType.Boolean:
                    Result = IntrinsicType.Boolean;
                    break;

                case TokenType.SByte:
                    Result = IntrinsicType.SByte;
                    break;

                case TokenType.Byte:
                    Result = IntrinsicType.Byte;
                    break;

                case TokenType.Short:
                    Result = IntrinsicType.Short;
                    break;

                case TokenType.UShort:
                    Result = IntrinsicType.UShort;
                    break;

                case TokenType.Integer:
                    Result = IntrinsicType.Integer;
                    break;

                case TokenType.UInteger:
                    Result = IntrinsicType.UInteger;
                    break;

                case TokenType.Long:
                    Result = IntrinsicType.Long;
                    break;

                case TokenType.ULong:
                    Result = IntrinsicType.ULong;
                    break;

                case TokenType.Decimal:
                    Result = IntrinsicType.Decimal;
                    break;

                case TokenType.Single:
                    Result = IntrinsicType.Single;
                    break;

                case TokenType.Double:
                    Result = IntrinsicType.Double;
                    break;

                case TokenType.Date:
                    Result = IntrinsicType.Date;
                    break;

                case TokenType.Char:
                    Result = IntrinsicType.Char;
                    break;

                case TokenType.String:
                    Result = IntrinsicType.String;
                    break;

                case TokenType.Object:
                    Result = IntrinsicType.Object;
                    break;

                case TokenType.Variant:
                    //TODO: Do not allow Variant in VB .NET
                    Result = IntrinsicType.Variant;
                    break;
                case TokenType.Currency:
                    //TODO: Do not allow Currency in VB .NET
                    Result = IntrinsicType.Currency;
                    break;
                default:
                    Result = IntrinsicType.None;
                    break;
            }

            return Result;
        }

        public override TokenType AsUnreservedKeyword()
        {
            return _UnreservedType;
        }
        #endregion

        #region Properties

        /// <summary>
        /// The identifier name.
        /// </summary>
        public string Identifier
        {
            get { return _Identifier; }
        }

        /// <summary>
        /// Whether the identifier is escaped.
        /// </summary>
        public bool Escaped
        {
            get { return _Escaped; }
        }

        /// <summary>
        /// The type character of the identifier.
        /// </summary>
        public TypeCharacter TypeCharacter
        {
            get { return _TypeCharacter; }
        }

        public bool HasSpaceAfter
        {
            get { return _hasSpaceAfter; }
        }

        #endregion

        #region Public Methods
        
        public override string ToString()
        {
            return this.Identifier;
        }

        #endregion

        #region Private Methods

        private static void BuildTable()
        {
            //// Unreserved Keywords: Can be used as identifier
            List<Keyword> KeywordList = new List<Keyword>();

            // NOTE: These have to be in the same order as the enum!
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.AddHandler));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.AddressOf));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Alias));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.And));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.AndAlso));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.VB7, TokenType.Ansi));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.As));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.VB7, TokenType.Assembly));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.VB7, TokenType.Auto));
            KeywordList.Add(new Keyword(LanguageVersion.VB6, LanguageVersion.Non, TokenType.Base));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.Non, TokenType.Binary));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Boolean));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.ByRef));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Byte));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.ByVal));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Call));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Case));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.Catch));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.CBool));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.CByte));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.CChar));
            KeywordList.Add(new Keyword(LanguageVersion.VB6, LanguageVersion.VB6, TokenType.CCur));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.CDate));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.CDbl));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.CDec));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.Char));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.CInt));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.Class));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.CLng));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.CObj));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.Non, TokenType.Compare));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Const));
            KeywordList.Add(new Keyword(LanguageVersion.VB8, LanguageVersion.VB8, TokenType.Continue));
            KeywordList.Add(new Keyword(LanguageVersion.VB8, LanguageVersion.VB8, TokenType.CSByte));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.CShort));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.CSng));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.CStr));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.CType));
            KeywordList.Add(new Keyword(LanguageVersion.VB8, LanguageVersion.VB8, TokenType.CUInt));
            KeywordList.Add(new Keyword(LanguageVersion.VB8, LanguageVersion.VB8, TokenType.CULng));
            KeywordList.Add(new Keyword(LanguageVersion.VB6, LanguageVersion.VB6, TokenType.Currency));
            KeywordList.Add(new Keyword(LanguageVersion.VB8, LanguageVersion.VB8, TokenType.CUShort));
            KeywordList.Add(new Keyword(LanguageVersion.VB6, LanguageVersion.VB6, TokenType.CVar));
            KeywordList.Add(new Keyword(LanguageVersion.VB8, LanguageVersion.Non, TokenType.Custom));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.Net, TokenType.Date));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Decimal));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Declare));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.Default));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.Delegate));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Dim));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.DirectCast));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Do));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Double));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Each));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Else));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.ElseIf));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.End));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.EndIf));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Enum));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Erase));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.Net, TokenType.Error));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Event));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Exit));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.Non, TokenType.Explicit));
            KeywordList.Add(new Keyword(LanguageVersion.VB8, LanguageVersion.Non, TokenType.ExternalChecksum));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Non, TokenType.ExternalSource));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.False));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.Finally));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.For));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Friend));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Function));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Get));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.GetType));
            KeywordList.Add(new Keyword(LanguageVersion.VB8 | LanguageVersion.VB6, LanguageVersion.VB8 | LanguageVersion.VB6, TokenType.Global));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.GoSub));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.GoTo));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.Handles));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.If));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Implements));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.Imports));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.In));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.Inherits));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Integer));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.Interface));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Is));
            KeywordList.Add(new Keyword(LanguageVersion.VB8, LanguageVersion.Non, TokenType.IsFalse));
            KeywordList.Add(new Keyword(LanguageVersion.VB8, LanguageVersion.VB8, TokenType.IsNot));
            KeywordList.Add(new Keyword(LanguageVersion.VB8, LanguageVersion.Non, TokenType.IsTrue));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Let));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Lib));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Like));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Long));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Loop));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Me));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.Mid));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Mod));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.Module));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.MustInherit));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.MustOverride));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.MyBase));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.MyClass));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.Namespace));
            KeywordList.Add(new Keyword(LanguageVersion.VB8, LanguageVersion.VB8, TokenType.Narrowing));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.New));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Next));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Not));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Nothing));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.NotInheritable));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.NotOverridable));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.Net, TokenType.Object));
            KeywordList.Add(new Keyword(LanguageVersion.VB8, LanguageVersion.VB8, TokenType.Of));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.Off));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.On));
            KeywordList.Add(new Keyword(LanguageVersion.VB8, LanguageVersion.VB8, TokenType.Operator));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Option));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Optional));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Or));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.OrElse));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.Overloads));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.Overridable));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.Overrides));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.ParamArray));
            KeywordList.Add(new Keyword(LanguageVersion.VB8, LanguageVersion.VB8, TokenType.Partial));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.VB7, TokenType.Preserve));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Private));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Property));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Protected));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Public));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.RaiseEvent));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.ReadOnly));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.ReDim));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Non, TokenType.Region));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Rem));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.RemoveHandler));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Resume));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Return));
            KeywordList.Add(new Keyword(LanguageVersion.VB8, LanguageVersion.VB8, TokenType.SByte));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Select));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Set));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.Shadows));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.Shared));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.Short));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Single));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Static));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Step));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Stop));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Non, TokenType.Strict));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.Net, TokenType.String));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.Structure));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Sub));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.SyncLock));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.Non, TokenType.Text));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Then));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.Throw));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.To));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.True));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.Try));
            KeywordList.Add(new Keyword(LanguageVersion.VB8, LanguageVersion.VB8, TokenType.TryCast));
            KeywordList.Add(new Keyword(LanguageVersion.VB6, LanguageVersion.VB6, TokenType.Type));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.TypeOf));
            KeywordList.Add(new Keyword(LanguageVersion.VB8, LanguageVersion.VB8, TokenType.UInteger));
            KeywordList.Add(new Keyword(LanguageVersion.VB8, LanguageVersion.VB8, TokenType.ULong));
            KeywordList.Add(new Keyword(LanguageVersion.VB8, LanguageVersion.VB8, TokenType.UShort));
            KeywordList.Add(new Keyword(LanguageVersion.VB8, LanguageVersion.VB8, TokenType.Using));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.VB7, TokenType.Unicode));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.VB7, TokenType.Until));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Variant));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Wend));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.When));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.While));
            KeywordList.Add(new Keyword(LanguageVersion.VB8, LanguageVersion.VB8, TokenType.Widening));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.With));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.WithEvents));
            KeywordList.Add(new Keyword(LanguageVersion.Net, LanguageVersion.Net, TokenType.WriteOnly));
            KeywordList.Add(new Keyword(LanguageVersion.All, LanguageVersion.All, TokenType.Xor));

            Dictionary<string, Keyword> Table = new Dictionary<string, Keyword>(StringComparer.InvariantCultureIgnoreCase);

            foreach (Keyword Item in KeywordList)
                AddKeyword(Table, Item.TokenType.ToString(), Item);

            KeywordTable = Table;
        }

        private static void AddKeyword(Dictionary<string, Keyword> table, string name, Keyword keyword)
        {
            table.Add(name, keyword);
            table.Add(Scanner.MakeFullWidth(name), keyword);
        }

        #endregion

    }
}

using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
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
using VBConverter.CodeParser.Tokens;
using VBConverter.CodeParser.Tokens.Literal;
using VBConverter.CodeParser.Tokens.Terminator;

namespace VBConverter.CodeParser.Serializers
{
    public class TokenXmlSerializer : XmlSerializer<Token>
    {
        #region Contructor

        public TokenXmlSerializer(XmlWriter Writer) : base(Writer)
        {
        }

        public TokenXmlSerializer(XmlWriter Writer, bool showPosition) : base(Writer, showPosition)
        {
        }

        #endregion

        #region Public Methods

        public override void Serialize(Token Token)
        {
            Output.WriteStartElement(Token.Type.ToString());
            SerializePosition(Token.Span);

            switch (Token.Type)
            {
                case TokenType.LexicalError:
                    Output.WriteAttributeString("errorNumber", ((ErrorToken)Token).SyntaxError.Type.ToString());
                    Output.WriteString(((ErrorToken)Token).SyntaxError.ToString());
                    break;

                case TokenType.Comment:
                    Output.WriteAttributeString("isRem", ((CommentToken)Token).IsREM.ToString());
                    Output.WriteString(((CommentToken)Token).Comment);
                    break;

                case TokenType.Identifier:
                    Output.WriteAttributeString("escaped", ((IdentifierToken)Token).Escaped.ToString());
                    Serialize(((IdentifierToken)Token).TypeCharacter);
                    Output.WriteString(((IdentifierToken)Token).Identifier);
                    break;

                case TokenType.StringLiteral:
                    Output.WriteString(((StringLiteralToken)Token).Literal);
                    break;

                case TokenType.CharacterLiteral:
                    Output.WriteString(((CharacterLiteralToken)Token).Literal.ToString());
                    break;

                case TokenType.DateLiteral:
                    Output.WriteString(((DateLiteralToken)Token).Literal.ToString());
                    break;

                case TokenType.IntegerLiteral:
                    Output.WriteAttributeString("base", ((IntegerLiteralToken)Token).IntegerBase.ToString());
                    Serialize(((IntegerLiteralToken)Token).TypeCharacter);
                    Output.WriteString(((IntegerLiteralToken)Token).Literal.ToString());
                    break;

                case TokenType.FloatingPointLiteral:
                    Serialize(((FloatingPointLiteralToken)Token).TypeCharacter);
                    Output.WriteString(((FloatingPointLiteralToken)Token).Literal.ToString());
                    break;

                case TokenType.DecimalLiteral:
                    Serialize(((DecimalLiteralToken)Token).TypeCharacter);
                    Output.WriteString(((DecimalLiteralToken)Token).Literal.ToString());
                    break;

                case TokenType.UnsignedIntegerLiteral:
                    Serialize(((UnsignedIntegerLiteralToken)Token).TypeCharacter);
                    Output.WriteString(((UnsignedIntegerLiteralToken)Token).Literal.ToString());
                    break;

                default:
                    break;
                // Fall through
            }

            Output.WriteEndElement();
        }

        #endregion

        #region Private Methods

        private void Serialize(TypeCharacter TypeCharacter)
        {
            if (TypeCharacter != TypeCharacter.None)
            {
                Dictionary<TypeCharacter, string> TypeCharacterTable = null;

                if (TypeCharacterTable == null)
                {
                    Dictionary<TypeCharacter, string> Table = new Dictionary<TypeCharacter, string>();
                    // NOTE: These have to be in the same order as the enum!
                    string[] TypeCharacters = { "$", "%", "&", "S", "I", "L", "!", "#", "@", "F", "R", "D", "US", "UI", "UL" };
                    TypeCharacter TableTypeCharacter = TypeCharacter.StringSymbol;

                    for (int Index = 0; Index <= TypeCharacters.Length - 1; Index++)
                    {
                        Table.Add(TableTypeCharacter, TypeCharacters[Index]);
                        TableTypeCharacter = (TypeCharacter)((int)TableTypeCharacter << 1);
                    }

                    TypeCharacterTable = Table;
                }

                Output.WriteAttributeString("typeChar", TypeCharacterTable[TypeCharacter]);
            }
        }

        #endregion
    }
}

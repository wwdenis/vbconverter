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
	/// A punctuation token.
	/// </summary>
	public sealed class PunctuatorToken : Token
    {
        #region Fields

        private static Dictionary<string, TokenType> PunctuatorTable;

        #endregion

        #region Constructors

        /// <summary>
        /// Constructs a new punctuator token.
        /// </summary>
        /// <param name="type">The punctuator token type.</param>
        /// <param name="span">The location of the punctuator.</param>
        public PunctuatorToken(TokenType type, Span span)
            : base(type, span)
        {
            if ((type < TokenType.LeftParenthesis || type > TokenType.GreaterThanGreaterThanEquals))
                throw new ArgumentOutOfRangeException("type");
        }

        static PunctuatorToken()
        {
            BuildTable();
        }

        #endregion

        #region Internal Methods

        // Returns the token type of the string.
		static internal TokenType Parse(string s)
		{
			if (!PunctuatorTable.ContainsKey(s))
				return TokenType.None;
			else
				return PunctuatorTable[s];
        }

        #endregion

        #region Private Methods

        private static void BuildTable()
        {
            Dictionary<string, TokenType> Table = new Dictionary<string, TokenType>(StringComparer.InvariantCulture);

            // NOTE: These have to be in the same order as the enum!
            AddPunctuator(Table, "(", TokenType.LeftParenthesis);
            AddPunctuator(Table, ")", TokenType.RightParenthesis);
            AddPunctuator(Table, "{", TokenType.LeftCurlyBrace);
            AddPunctuator(Table, "}", TokenType.RightCurlyBrace);
            AddPunctuator(Table, "!", TokenType.Exclamation);
            AddPunctuator(Table, "#", TokenType.Pound);
            AddPunctuator(Table, ",", TokenType.Comma);
            AddPunctuator(Table, ".", TokenType.Period);
            AddPunctuator(Table, ":", TokenType.Colon);
            AddPunctuator(Table, ":=", TokenType.ColonEquals);
            AddPunctuator(Table, "&", TokenType.Ampersand);
            AddPunctuator(Table, "&=", TokenType.AmpersandEquals);
            AddPunctuator(Table, "*", TokenType.Star);
            AddPunctuator(Table, "*=", TokenType.StarEquals);
            AddPunctuator(Table, "+", TokenType.Plus);
            AddPunctuator(Table, "+=", TokenType.PlusEquals);
            AddPunctuator(Table, "-", TokenType.Minus);
            AddPunctuator(Table, "-=", TokenType.MinusEquals);
            AddPunctuator(Table, "/", TokenType.ForwardSlash);
            AddPunctuator(Table, "/=", TokenType.ForwardSlashEquals);
            AddPunctuator(Table, "\\", TokenType.BackwardSlash);
            AddPunctuator(Table, "\\=", TokenType.BackwardSlashEquals);
            AddPunctuator(Table, "^", TokenType.Caret);
            AddPunctuator(Table, "^=", TokenType.CaretEquals);
            AddPunctuator(Table, "<", TokenType.LessThan);
            AddPunctuator(Table, "<=", TokenType.LessThanEquals);
            AddPunctuator(Table, "=", TokenType.Equals);
            AddPunctuator(Table, "<>", TokenType.NotEquals);
            AddPunctuator(Table, ">", TokenType.GreaterThan);
            AddPunctuator(Table, ">=", TokenType.GreaterThanEquals);
            AddPunctuator(Table, "<<", TokenType.LessThanLessThan);
            AddPunctuator(Table, "<<=", TokenType.LessThanLessThanEquals);
            AddPunctuator(Table, ">>", TokenType.GreaterThanGreaterThan);
            AddPunctuator(Table, ">>=", TokenType.GreaterThanGreaterThanEquals);

            PunctuatorTable = Table;
        }

        private static void AddPunctuator(Dictionary<string, TokenType> table, string punctuator, TokenType type)
        {
            table.Add(punctuator, type);
            table.Add(Scanner.MakeFullWidth(punctuator), type);
        }

        #endregion
    }
}

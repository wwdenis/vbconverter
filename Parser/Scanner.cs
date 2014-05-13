using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
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
using VBConverter.CodeParser.Error;
using VBConverter.CodeParser.Tokens;
using VBConverter.CodeParser.Tokens.Literal;
using VBConverter.CodeParser.Tokens.Terminator;

namespace VBConverter.CodeParser
{
    /// <summary>
    /// A lexical analyzer for Visual Basic .NET. It produces a stream of lexical tokens.
    /// </summary>
    public sealed class Scanner : IDisposable
    {

        #region Fields

        // Version of the language to parse
        private LanguageVersion _Version = LanguageVersion.VB8;

        // The text to be read. We use a TextReader here so that lexical analysis
        // can be done on strings as well as streams.
        private TextReader _Source;

        // How many columns a tab character should be treated as
        private int _TabSpaces = 4;

        // For performance reasons, we cache the character when we peek ahead.
        private char _PeekCache;
        private bool _PeekCacheHasValue = false;

        // There are a few places where we're going to need to peek one character
        // ahead
        private char _PeekAheadCache;
        private bool _PeekAheadCacheHasValue = false;

        // Since we're only using a TextReader which has no position information,
        // we have to keep track of line/column information ourselves.
        private long _Index = 0;
        private long _Line = 1;
        private long _Column = 1;

        // A buffer of all the tokens we've returned so far
        private List<Token> _Tokens = new List<Token>();

        // Our current position in the buffer. -1 means before the beginning.
        private int _Position = -1;

        private int _CurrentPosition = -1;

        private bool _IsLocked = false; 

        // Determine whether we have been disposed already or not
        private bool _Disposed = false;

        #endregion

        #region Properties

        /// <summary>
        /// The version of Visual Basic this scanner operates on.
        /// </summary>
        public LanguageVersion Version
        {
            get { return _Version; }
            set
            {
                if (!Enum.IsDefined(typeof(LanguageVersion), value))
                    throw new ArgumentOutOfRangeException("Version");

                _Version = value;
            }
        }

        /// <summary>
        /// TextReader object that representes the text to be read.
        /// </summary>
        private TextReader Source
        {
            get 
            { 
                return _Source; 
            }
            set
            {
                if (value == null)
                    throw new ArgumentNullException("Source");

                _Source = value;
            }
        }

        /// <summary>
        /// How many columns a tab character should be considered.
        /// </summary>
        public int TabSpaces
        {
            get { return _TabSpaces; }

            set
            {
                if (value < 1)
                    throw new ArgumentException("Tabs cannot represent less than one space.");
                
                _TabSpaces = value;
            }
        }

        /// <summary>
        /// Whether the stream is positioned on the first token.
        /// </summary>
        public bool IsOnFirstToken
        {
            get { return Position == -1; }
        }

        public bool IsEndOfFile
        {
            get { return Peek() != null && Peek().Type == TokenType.EndOfStream; }
        }

        public int Position
        {
            get 
            { 
                return _Position; 
            }
            private set 
            { 
                _Position = value;

                if (!IsLocked)
                    _CurrentPosition = _Position;
            }
        }

        public bool IsLocked
        {
            get { return _IsLocked; }
            set 
            { 
                _IsLocked = value;
                Position = _Position;
            }
        }

        public Token Current
        {
            get 
            {
                int NewPosition = _CurrentPosition + 1;
                if (_CurrentPosition >= 0)
                    return _Tokens[_CurrentPosition];
                else
                    return null;
            }
        }

        /// <summary>
        /// The current line/column position
        /// </summary>
        private Location CurrentLocation
        {
            get { return new Location(_Index, _Line, _Column); }
        }



        #endregion

        #region Constructors

        /// <summary>
        /// Constructs a scanner for a general TextReader.
        /// </summary>
        /// <param name="source">The TextReader to scan.</param>
        public Scanner(TextReader source)
        {
            this.Source = source;
        }

        /// <summary>
        /// Constructs a scanner for a stream.
        /// </summary>
        /// <param name="source">The stream to scan.</param>
        public Scanner(Stream source)
        {
            this.Source = new StreamReader(source);
        }

        /// <summary>
        /// Constructs a scanner for a string.
        /// </summary>
        /// <param name="source">The string to scan.</param>
        public Scanner(string source)
        {
            this.Source = new StringReader(source);
        }

        /// <summary>
        /// Constructs a canner for a general TextReader.
        /// </summary>
        /// <param name="source">The TextReader to scan.</param>
        /// <param name="version">The language version to parse.</param>
        public Scanner(TextReader source, LanguageVersion version) : this(source)
        {
            this.Version = version;
        }

        /// <summary>
        /// Constructs a scanner for a stream.
        /// </summary>
        /// <param name="source">The stream to scan.</param>
        /// <param name="version">The language version to parse.</param>
        public Scanner(Stream source, LanguageVersion version) : this(source)
        {
            this.Version = version;
        }

        /// <summary>
        /// Constructs a scanner for a string.
        /// </summary>
        /// <param name="source">The string to scan.</param>
        /// <param name="version">The language version to parse.</param>
        public Scanner(string source, LanguageVersion version) : this(source)
        {
            this.Version = version;
        }

        #endregion

        #region Private Methods

        // Read a character
        private char ReadChar()
        {
            char c;

            if (_PeekCacheHasValue)
            {
                c = _PeekCache;
                _PeekCacheHasValue = false;

                if (_PeekAheadCacheHasValue)
                {
                    _PeekCache = _PeekAheadCache;
                    _PeekCacheHasValue = true;
                    _PeekAheadCacheHasValue = false;
                }
            }
            else
            {
                Debug.Assert(!_PeekAheadCacheHasValue, "Cache incorrect!");
                c = (char)(this.Source.Read());
            }

            _Index += 1;
            if ((int)(c) == 9)
                _Column += _TabSpaces;
            else
                _Column += 1;
            
            return c;
        }

        // Peek ahead at the next character
        private char PeekChar()
        {
            if (!_PeekCacheHasValue)
            {
                _PeekCache = (char)(this.Source.Read());
                _PeekCacheHasValue = true;
            }

            return _PeekCache;
        }

        // Peek at the character past the next character
        private char PeekAheadChar()
        {
            if (!_PeekAheadCacheHasValue)
            {
                if (!_PeekCacheHasValue)
                {
                    PeekChar();
                }

                _PeekAheadCache = (char)(this.Source.Read());
                _PeekAheadCacheHasValue = true;
            }

            return _PeekAheadCache;
        }

        // Creates a span from the start location to the current location.
        private Span SpanFrom(Location start)
        {
            return new Span(start, CurrentLocation);
        }



        //
        // Scan functions
        //
        // Each function assumes that the reader is positioned at the beginning of
        // the token. At the end, the function will have read through the entire
        // token. If an error occurs, the function may attempt to do error recovery.
        //

        private TypeCharacter ScanPossibleTypeCharacter(TypeCharacter ValidTypeCharacters)
        {
            char TypeChar = PeekChar();
            string TypeString;

            if (static_ScanPossibleTypeCharacter_TypeCharacterTable == null)
            {
                Dictionary<string, TypeCharacter> Table = new Dictionary<string, TypeCharacter>(StringComparer.InvariantCultureIgnoreCase);
                // NOTE: These have to be in the same order as the enum!
                string[] TypeCharacters = {"$", "%", "&", "S", "I", "L", "!", "#", "@", "F",  "R", "D", "US", "UI", "UL"};
                TypeCharacter Character = TypeCharacter.StringSymbol;

                for (int Index = 0; Index <= TypeCharacters.Length - 1; Index++)
                {
                    Table.Add(TypeCharacters[Index], Character);
                    Table.Add(Scanner.MakeFullWidth(TypeCharacters[Index]), Character);
                    
                    Character = (TypeCharacter)((int)Character << 1);
                }

                static_ScanPossibleTypeCharacter_TypeCharacterTable = Table;
            }

            if (IsUnsignedTypeChar(TypeChar) && LanguageInfo.ImplementsVB80(_Version))
            {
                // At the point at which we've seen a "U", we don't know if it's going to
                // be a valid type character or just something invalid.
                TypeString = TypeChar.ToString() + PeekAheadChar().ToString();
            }
            else
            {
                TypeString = TypeChar.ToString();
            }

            if (static_ScanPossibleTypeCharacter_TypeCharacterTable.ContainsKey(TypeString))
            {
                TypeCharacter Character = static_ScanPossibleTypeCharacter_TypeCharacterTable[TypeString];

                if ((Character & ValidTypeCharacters) != 0)
                {
                    // A bang (!) is a type character unless it is followed by a legal identifier start.
                    if (Character == TypeCharacter.SingleSymbol && CanStartIdentifier(PeekAheadChar()))
                    {
                        return TypeCharacter.None;
                    }

                    ReadChar();

                    if (IsUnsignedTypeChar(TypeChar))
                    {
                        ReadChar();
                    }

                    return Character;
                }
            }

            return TypeCharacter.None;
        }
        static Dictionary<string, TypeCharacter> static_ScanPossibleTypeCharacter_TypeCharacterTable;

        private PunctuatorToken ScanPossibleMultiCharacterPunctuator(char leadingCharacter, Location start)
        {
            char NextChar = PeekChar();
            TokenType Punctuator;
            string PunctuatorString = leadingCharacter.ToString();

            Debug.Assert(PunctuatorToken.Parse(leadingCharacter.ToString()) != TokenType.None);

            if (IsEquals(NextChar) || IsLessThan(NextChar) || IsGreaterThan(NextChar))
            {
                PunctuatorString = PunctuatorString + NextChar;
                Punctuator = PunctuatorToken.Parse(PunctuatorString);

                if (Punctuator != TokenType.None)
                {
                    ReadChar();

                    if ((Punctuator == TokenType.LessThanLessThan || Punctuator == TokenType.GreaterThanGreaterThan) && IsEquals(PeekChar()))
                    {
                        PunctuatorString = PunctuatorString + ReadChar();
                        Punctuator = PunctuatorToken.Parse(PunctuatorString);
                    }

                    return new PunctuatorToken(Punctuator, SpanFrom(start));
                }
            }

            Punctuator = PunctuatorToken.Parse(leadingCharacter.ToString());
            return new PunctuatorToken(Punctuator, SpanFrom(start));
        }

        private Token ScanNumericLiteral()
        {
            Location Start = CurrentLocation;
            StringBuilder Literal = new StringBuilder();
            IntegerBase Base = IntegerBase.Decimal;
            TypeCharacter TypeCharacter = TypeCharacter.None;

            Debug.Assert(CanStartNumericLiteral());

            if (IsAmpersand(PeekChar()))
            {
                Literal.Append(MakeHalfWidth(ReadChar()));

                if (IsHexDesignator(PeekChar()))
                {
                    Literal.Append(MakeHalfWidth(ReadChar()));
                    Base = IntegerBase.Hexadecimal;

                    while (IsHexDigit(PeekChar()))
                    {
                        Literal.Append(MakeHalfWidth(ReadChar()));
                    }
                }
                else if (IsOctalDesignator(PeekChar()))
                {
                    Literal.Append(MakeHalfWidth(ReadChar()));
                    Base = IntegerBase.Octal;

                    while (IsOctalDigit(PeekChar()))
                    {
                        Literal.Append(MakeHalfWidth(ReadChar()));
                    }
                }
                else
                {
                    return ScanPossibleMultiCharacterPunctuator('&', Start);
                }

                if (Literal.Length > 2)
                {
                    const TypeCharacter ValidTypeChars = TypeCharacter.ShortChar | TypeCharacter.UnsignedShortChar | TypeCharacter.IntegerSymbol | TypeCharacter.IntegerChar | TypeCharacter.UnsignedIntegerChar | TypeCharacter.LongSymbol | TypeCharacter.LongChar | TypeCharacter.UnsignedLongChar;

                    TypeCharacter = ScanPossibleTypeCharacter(ValidTypeChars);

                    try
                    {
                        long LongValue;
                        ulong ULongValue;
                        string StringValue = Literal.ToString();
                        NumberStyles Style = NumberStyles.Number;

                        if (Base == IntegerBase.Hexadecimal)
                        {
                            Style = NumberStyles.HexNumber;
                            StringValue = StringValue.ToUpper();
                            StringValue = StringValue.Replace("&H", "");
                        }

                        LongValue = long.Parse(StringValue, Style);
                        ULongValue = ulong.Parse(StringValue, Style);

                        switch (TypeCharacter)
                        {
                            case TypeCharacter.ShortChar:
                                if (LongValue <= 65535L)
                                {
                                    if (LongValue > 32767L)
                                        LongValue = -(65536L - LongValue);

                                    if (LongValue >= short.MinValue && LongValue <= short.MaxValue)
                                        return new IntegerLiteralToken((short)LongValue, Base, TypeCharacter, SpanFrom(Start));
                                }

                                break;
                            // Fall through

                            case TypeCharacter.UnsignedShortChar:
                                if (ULongValue <= 65535L)
                                {
                                    if (ULongValue >= ushort.MinValue && ULongValue <= ushort.MaxValue)
                                        return new UnsignedIntegerLiteralToken((ushort)ULongValue, Base, TypeCharacter, SpanFrom(Start));
                                }

                                break;
                            // Fall through

                            case TypeCharacter.IntegerSymbol:
                            case TypeCharacter.IntegerChar:
                                if (LongValue <= 4294967295L)
                                {
                                    if (LongValue > 2147483647L)
                                        LongValue = -(4294967296L - LongValue);

                                    if (LongValue >= int.MinValue && LongValue <= int.MaxValue)
                                        return new IntegerLiteralToken((int)LongValue, Base, TypeCharacter, SpanFrom(Start));
                                }

                                break;
                            // Fall through

                            case TypeCharacter.UnsignedIntegerChar:
                                if (ULongValue <= 4294967295L)
                                {
                                    if (ULongValue >= uint.MinValue && ULongValue <= uint.MaxValue)
                                        return new UnsignedIntegerLiteralToken((uint)ULongValue, Base, TypeCharacter, SpanFrom(Start));
                                }

                                break;
                            // Fall through

                            case TypeCharacter.LongSymbol:
                            case TypeCharacter.LongChar:
                                return new IntegerLiteralToken(LongValue, Base, TypeCharacter, SpanFrom(Start));

                            case TypeCharacter.UnsignedLongChar:

                                return new UnsignedIntegerLiteralToken(ULongValue, Base, TypeCharacter, SpanFrom(Start));

                            default:
                                TypeCharacter = TypeCharacter.None;
                                return new IntegerLiteralToken(LongValue, Base, TypeCharacter, SpanFrom(Start));
                        }
                    }
                    catch (OverflowException)
                    {
                        return new ErrorToken(SyntaxErrorType.InvalidIntegerLiteral, SpanFrom(Start));
                    }
                    catch (InvalidCastException)
                    {
                        return new ErrorToken(SyntaxErrorType.InvalidIntegerLiteral, SpanFrom(Start));
                    }
                }

                return new ErrorToken(SyntaxErrorType.InvalidIntegerLiteral, SpanFrom(Start));
            }

            while (IsDigit(PeekChar()))
            {
                Literal.Append(MakeHalfWidth(ReadChar()));
            }

            if (IsPeriod(PeekChar()) || IsExponentDesignator(PeekChar()))
            {
                SyntaxErrorType ErrorType = SyntaxErrorType.None;
                const TypeCharacter ValidTypeChars = TypeCharacter.DecimalChar | TypeCharacter.DecimalSymbol | TypeCharacter.SingleChar | TypeCharacter.SingleSymbol | TypeCharacter.DoubleChar | TypeCharacter.DoubleSymbol;

                if (IsPeriod(PeekChar()))
                {
                    Literal.Append(MakeHalfWidth(ReadChar()));

                    if (!IsDigit(PeekChar()) & Literal.Length == 1)
                    {
                        return new PunctuatorToken(TokenType.Period, SpanFrom(Start));
                    }

                    while (IsDigit(PeekChar()))
                    {
                        Literal.Append(MakeHalfWidth(ReadChar()));
                    }
                }

                if (IsExponentDesignator(PeekChar()))
                {
                    Literal.Append(MakeHalfWidth(ReadChar()));

                    if (IsPlus(PeekChar()) || IsMinus(PeekChar()))
                    {
                        Literal.Append(MakeHalfWidth(ReadChar()));
                    }

                    if (!IsDigit(PeekChar()))
                    {
                        return new ErrorToken(SyntaxErrorType.InvalidFloatingPointLiteral, SpanFrom(Start));
                    }

                    while (IsDigit(PeekChar()))
                    {
                        Literal.Append(MakeHalfWidth(ReadChar()));
                    }
                }

                TypeCharacter = ScanPossibleTypeCharacter(ValidTypeChars);

                try
                {
                    switch (TypeCharacter)
                    {
                        case TypeCharacter.DecimalChar:
                        case TypeCharacter.DecimalSymbol:
                            ErrorType = SyntaxErrorType.InvalidDecimalLiteral;
                            return new DecimalLiteralToken(Convert.ToDecimal(Literal.ToString()), TypeCharacter, SpanFrom(Start));

                        case TypeCharacter.SingleSymbol:
                        case TypeCharacter.SingleChar:
                            ErrorType = SyntaxErrorType.InvalidFloatingPointLiteral;
                            return new FloatingPointLiteralToken(Convert.ToSingle(Literal.ToString()), TypeCharacter, SpanFrom(Start));

                        case TypeCharacter.DoubleSymbol:
                        case TypeCharacter.DoubleChar:
                            ErrorType = SyntaxErrorType.InvalidFloatingPointLiteral;
                            return new FloatingPointLiteralToken(Convert.ToDouble(Literal.ToString()), TypeCharacter, SpanFrom(Start));

                        default:
                            ErrorType = SyntaxErrorType.InvalidFloatingPointLiteral;
                            TypeCharacter = TypeCharacter.None;
                            return new FloatingPointLiteralToken(Convert.ToDouble(Literal.ToString()), TypeCharacter, SpanFrom(Start));
                    }
                }
                catch (OverflowException)
                {
                    return new ErrorToken(ErrorType, SpanFrom(Start));
                }
                catch (InvalidCastException)
                {
                    return new ErrorToken(ErrorType, SpanFrom(Start));
                }
            }
            else
            {
                SyntaxErrorType ErrorType = SyntaxErrorType.None;
                const TypeCharacter ValidTypeChars = TypeCharacter.ShortChar | TypeCharacter.IntegerSymbol | TypeCharacter.IntegerChar | TypeCharacter.LongSymbol | TypeCharacter.LongChar | TypeCharacter.DecimalSymbol | TypeCharacter.DecimalChar | TypeCharacter.SingleSymbol | TypeCharacter.SingleChar | TypeCharacter.DoubleSymbol | TypeCharacter.DoubleChar | TypeCharacter.UnsignedShortChar | TypeCharacter.UnsignedIntegerChar | TypeCharacter.UnsignedLongChar;

                TypeCharacter = ScanPossibleTypeCharacter(ValidTypeChars);

                try
                {
                    switch (TypeCharacter)
                    {
                        case TypeCharacter.ShortChar:
                            ErrorType = SyntaxErrorType.InvalidIntegerLiteral;
                            return new IntegerLiteralToken(Convert.ToInt16(Literal.ToString()), Base, TypeCharacter, SpanFrom(Start));

                        case TypeCharacter.UnsignedShortChar:
                            ErrorType = SyntaxErrorType.InvalidIntegerLiteral;
                            return new UnsignedIntegerLiteralToken(Convert.ToUInt16(Literal.ToString()), Base, TypeCharacter, SpanFrom(Start));

                        case TypeCharacter.IntegerSymbol:
                        case TypeCharacter.IntegerChar:
                            ErrorType = SyntaxErrorType.InvalidIntegerLiteral;
                            return new IntegerLiteralToken(Convert.ToInt32(Literal.ToString()), Base, TypeCharacter, SpanFrom(Start));

                        case TypeCharacter.UnsignedIntegerChar:
                            ErrorType = SyntaxErrorType.InvalidIntegerLiteral;
                            return new UnsignedIntegerLiteralToken(Convert.ToUInt32(Literal.ToString()), Base, TypeCharacter, SpanFrom(Start));

                        case TypeCharacter.LongSymbol:
                        case TypeCharacter.LongChar:
                            ErrorType = SyntaxErrorType.InvalidIntegerLiteral;
                            return new IntegerLiteralToken(Convert.ToInt64(Literal.ToString()), Base, TypeCharacter, SpanFrom(Start));

                        case TypeCharacter.UnsignedLongChar:
                            ErrorType = SyntaxErrorType.InvalidIntegerLiteral;
                            return new UnsignedIntegerLiteralToken(Convert.ToUInt64(Literal.ToString()), Base, TypeCharacter, SpanFrom(Start));

                        case TypeCharacter.DecimalChar:
                        case TypeCharacter.DecimalSymbol:
                            ErrorType = SyntaxErrorType.InvalidDecimalLiteral;
                            return new DecimalLiteralToken(Convert.ToDecimal(Literal.ToString()), TypeCharacter, SpanFrom(Start));

                        case TypeCharacter.SingleSymbol:
                        case TypeCharacter.SingleChar:
                            ErrorType = SyntaxErrorType.InvalidFloatingPointLiteral;
                            return new FloatingPointLiteralToken(Convert.ToDouble(Literal.ToString()), TypeCharacter, SpanFrom(Start));

                        case TypeCharacter.DoubleSymbol:
                        case TypeCharacter.DoubleChar:
                            ErrorType = SyntaxErrorType.InvalidFloatingPointLiteral;
                            return new FloatingPointLiteralToken(Convert.ToDouble(Literal.ToString()), TypeCharacter, SpanFrom(Start));

                        default:
                            ErrorType = SyntaxErrorType.InvalidIntegerLiteral;
                            return new IntegerLiteralToken(Convert.ToInt64(Literal.ToString()), Base, TypeCharacter.None, SpanFrom(Start));
                    }
                }
                catch (OverflowException)
                {
                    return new ErrorToken(ErrorType, SpanFrom(Start));
                }
                catch (InvalidCastException)
                {
                    return new ErrorToken(ErrorType, SpanFrom(Start));
                }
            }
        }

        private bool CanStartNumericLiteral()
        {
            return IsPeriod(PeekChar()) || IsAmpersand(PeekChar()) || IsDigit(PeekChar());
        }

        private long ReadIntegerLiteral()
        {
            long Value = 0;

            while (IsDigit(PeekChar()))
            {
                char c = MakeHalfWidth(ReadChar());
                Value *= 10;
                Value += (int)(c) - (int)('0');
            }

            return Value;
        }

        private Token ScanDateLiteral()
        {
            Location Start = CurrentLocation;
            Location PossibleEnd;
            int Month = 0;
            int Day = 0;
            int Year = 0;
            int Hour = 0;
            int Minute = 0;
            int Second = 0;
            bool HaveDateValue = false;
            bool HaveTimeValue = false;
            long Value;

            Debug.Assert(CanStartDateLiteral());

            ReadChar();
            PossibleEnd = CurrentLocation;
            EatWhitespace();

            // Distinguish between date literals and the # punctuator
            if (!IsDigit(PeekChar()))
            {
                return new PunctuatorToken(TokenType.Pound, new Span(Start, PossibleEnd));
            }

            Value = ReadIntegerLiteral();

            if (IsForwardSlash(PeekChar()) || IsMinus(PeekChar()))
            {
                char Separator = ReadChar();
                Location YearStart;

                HaveDateValue = true;
                if (Value < 1 || Value > 12) return GetInvalidDateLiteral(Start);
                Month = (int)Value;

                if (!IsDigit(PeekChar())) return GetInvalidDateLiteral(Start);
                Value = ReadIntegerLiteral();
                if (Value < 1 || Value > 31) return GetInvalidDateLiteral(Start);
                Day = (int)Value;

                if (PeekChar() != Separator) return GetInvalidDateLiteral(Start);
                ReadChar();

                if (!IsDigit(PeekChar())) return GetInvalidDateLiteral(Start);
                YearStart = CurrentLocation;
                Value = ReadIntegerLiteral();
                if (Value < 1 || Value > 9999) return GetInvalidDateLiteral(Start);
                // Years less than 1000 have to be four digits long to avoid y2k confusion
                if (Value < 1000 & CurrentLocation.Column - YearStart.Column != 4) return GetInvalidDateLiteral(Start);
                Year = (int)Value;

                if (Day > System.DateTime.DaysInMonth(Year, Month)) return GetInvalidDateLiteral(Start);

                EatWhitespace();
                if (IsDigit(PeekChar()))
                {
                    Value = ReadIntegerLiteral();

                    if (!IsColon(PeekChar())) return GetInvalidDateLiteral(Start);
                }
            }

            if (IsColon(PeekChar()))
            {
                ReadChar();
                HaveTimeValue = true;
                if (Value < 0 || Value > 23) return GetInvalidDateLiteral(Start);
                Hour = (int)Value;

                if (!IsDigit(PeekChar())) return GetInvalidDateLiteral(Start);
                Value = ReadIntegerLiteral();
                if (Value < 0 || Value > 59) return GetInvalidDateLiteral(Start);
                Minute = (int)Value;

                if (IsColon(PeekChar()))
                {
                    ReadChar();
                    if (!IsDigit(PeekChar())) return GetInvalidDateLiteral(Start);
                    Value = ReadIntegerLiteral();
                    if (Value < 0 || Value > 59) return GetInvalidDateLiteral(Start);
                    Second = (int)Value;
                }

                EatWhitespace();

                if (IsA(PeekChar()))
                {
                    ReadChar();

                    if (IsM(PeekChar()))
                    {
                        ReadChar();
                        if (Hour < 1 || Hour > 12)
                        {
                            return GetInvalidDateLiteral(Start);
                        }
                    }
                    else
                    {
                        return GetInvalidDateLiteral(Start);
                    }
                }
                else if (IsP(PeekChar()))
                {
                    ReadChar();

                    if (IsM(PeekChar()))
                    {
                        ReadChar();
                        if (Hour < 1 || Hour > 12)
                        {
                            return GetInvalidDateLiteral(Start);
                        }

                        Hour += 12;

                        if (Hour == 24)
                        {
                            Hour = 12;
                        }
                    }
                    else
                    {
                        return GetInvalidDateLiteral(Start);
                    }
                }
            }

            if (!IsPound(PeekChar()))
            {
                return GetInvalidDateLiteral(Start);
            }
            else
            {
                ReadChar();
            }

            if (!HaveTimeValue && !HaveDateValue)
            {
                return GetInvalidDateLiteral(Start);
            }

            if (HaveDateValue)
            {
                if (HaveTimeValue)
                {
                    return new DateLiteralToken(new System.DateTime(Year, Month, Day, Hour, Minute, Second), SpanFrom(Start));
                }
                else
                {
                    return new DateLiteralToken(new System.DateTime(Year, Month, Day), SpanFrom(Start));
                }
            }
            else
            {
                return new DateLiteralToken(new System.DateTime(1, 1, 1, Hour, Minute, Second), SpanFrom(Start));
            }
        }

        private ErrorToken GetInvalidDateLiteral(Location Start)
        {
            while (!IsPound(PeekChar()) && !CanStartLineTerminator())
            {
                ReadChar();
            }

            if (IsPound(PeekChar()))
            {
                ReadChar();
            }

            return new ErrorToken(SyntaxErrorType.InvalidDateLiteral, SpanFrom(Start));
        }

        private bool CanStartDateLiteral()
        {
            return IsPound(PeekChar());
        }

        // Actually, this scans string and char literals
        private Token ScanStringLiteral()
        {
            Location Start = CurrentLocation;
            StringBuilder s = new StringBuilder();

            Debug.Assert(CanStartStringLiteral());

            ReadChar();
        ContinueScan:

            while (!IsDoubleQuote(PeekChar()) && !CanStartLineTerminator())
            {
                s.Append(ReadChar());
            }

            if (IsDoubleQuote(PeekChar()))
            {
                ReadChar();

                if (IsDoubleQuote(PeekChar()))
                {
                    ReadChar();
                    // NOTE: We take what could be a full-width double quote and make it a half-width.
                    s.Append('"');
                    goto ContinueScan;
                }
            }
            else
            {
                return new ErrorToken(SyntaxErrorType.InvalidStringLiteral, SpanFrom(Start));
            }

            if (IsCharDesignator(PeekChar()))
            {
                ReadChar();

                if (s.Length != 1)
                {
                    return new ErrorToken(SyntaxErrorType.InvalidCharacterLiteral, SpanFrom(Start));
                }
                else
                {
                    return new CharacterLiteralToken(s[0], SpanFrom(Start));
                }
            }
            else
            {
                return new StringLiteralToken(s.ToString(), SpanFrom(Start));
            }
        }

        private bool CanStartStringLiteral()
        {
            return IsDoubleQuote(PeekChar());
        }

        private Token ScanIdentifier()
        {
            Location Start = CurrentLocation;
            bool Escaped = false;
            TypeCharacter TypeCharacter = TypeCharacter.None;
            string Identifier;
            StringBuilder s = new StringBuilder();
            TokenType Type = TokenType.Identifier;
            TokenType UnreservedType = TokenType.Identifier;

            Debug.Assert(CanStartIdentifier());

            if (IsLeftBracket(PeekChar()))
            {
                Escaped = true;
                ReadChar();

                if (!CanStartNonEscapedIdentifier())
                {
                    while (!IsRightBracket(PeekChar()) && !CanStartLineTerminator())
                        ReadChar();

                    if (IsRightBracket(PeekChar()))
                        ReadChar();

                    return new ErrorToken(SyntaxErrorType.InvalidEscapedIdentifier, SpanFrom(Start));
                }
            }

            s.Append(ReadChar());

            if (IsUnderscore(s[0]) && !IsIdentifier(PeekChar()))
            {
                Location End = CurrentLocation;

                EatWhitespace();

                // This is a line continuation
                if (CanStartLineTerminator())
                {
                    ScanLineTerminator(false);
                    return null;
                }
                else
                {
                    return new ErrorToken(SyntaxErrorType.InvalidIdentifier, new Span(Start, End));
                }
            }

            while (IsIdentifier(PeekChar()))
            {
                // NOTE: We do not convert full-width to half-width here!
                s.Append(ReadChar());
            }

            Identifier = s.ToString();

            if (Escaped)
            {
                if (IsRightBracket(PeekChar()))
                {
                    ReadChar();
                }
                else
                {
                    while (!IsRightBracket(PeekChar()) && !CanStartLineTerminator())
                        ReadChar();

                    if (IsRightBracket(PeekChar()))
                        ReadChar();

                    return new ErrorToken(SyntaxErrorType.InvalidEscapedIdentifier, SpanFrom(Start));
                }
            }
            else
            {
                const TypeCharacter ValidTypeChars = TypeCharacter.DecimalSymbol | TypeCharacter.DoubleSymbol | TypeCharacter.IntegerSymbol | TypeCharacter.LongSymbol | TypeCharacter.SingleSymbol | TypeCharacter.StringSymbol;

                Type = IdentifierToken.Parse(Identifier, _Version, false);

                if (Type == TokenType.Rem)
                    return ScanComment(Start);

                UnreservedType = IdentifierToken.Parse(Identifier, _Version, true);
                TypeCharacter = ScanPossibleTypeCharacter(ValidTypeChars);

                if (Type != TokenType.Identifier && TypeCharacter != TypeCharacter.None)
                {
                    // In VB 8.0, keywords with a type character are considered identifiers.
                    if (LanguageInfo.ImplementsVB80(_Version))
                        Type = TokenType.Identifier;
                    else
                        return new ErrorToken(SyntaxErrorType.InvalidTypeCharacterOnKeyword, SpanFrom(Start));
                }
            }
            
            char nextChar = PeekChar();
            bool hasSpaceAfter = (nextChar == ' ');

            return new IdentifierToken(Type, UnreservedType, Identifier, Escaped, hasSpaceAfter, TypeCharacter, SpanFrom(Start));
        }

        private bool CanStartNonEscapedIdentifier()
        {
            return CanStartNonEscapedIdentifier(PeekChar());
        }

        private bool CanStartIdentifier()
        {
            return CanStartIdentifier(PeekChar());
        }

        // Scan a comment that begins with a tick mark
        private CommentToken ScanComment()
        {
            return ScanComment(Location.Empty);
        }

        // Scan a comment that begins with REM.
        private CommentToken ScanComment(Location Start)
        {
            StringBuilder Builder = new StringBuilder();
            bool IsRem = true;
            
            if (Start == Location.Empty)
            {
                Debug.Assert(CanStartComment());

                IsRem = false;
                Start = CurrentLocation;

                if (PeekChar() == '\'')
                    ReadChar();
            }

            while (!CanStartLineTerminator())
            {
                char Item = ReadChar();
                // NOTE: We don't convert full-width to half-width here.
                Builder.Append(Item);
            }

            Span StartSpam = SpanFrom(Start);
            string CommentText = Builder.ToString();

            return new CommentToken(CommentText, IsRem, StartSpam);
        }

        // We only check for tick mark here.
        private bool CanStartComment()
        {
            char Current = PeekChar();
            bool IsQuote = IsSingleQuote(Current);
            bool IsContinuation = CanContinueComment();

            return IsQuote || IsContinuation;
        }

        private bool CanContinueComment()
        {
            if (!LanguageInfo.ImplementsVB60(Version))
                return false;

            bool Result = false;
            Token PreviousToken = null;
            if (Position > 0)
                PreviousToken = _Tokens[Position - 1];

            if (PreviousToken != null && PreviousToken.Type == TokenType.Comment)
            {
                string CommentText = ((CommentToken)PreviousToken).Comment;
                Result = CommentText.EndsWith("_");
            }

            return Result;
        }

        private Token ScanLineTerminator()
        {
            return ScanLineTerminator(true);
        }

        private Token ScanLineTerminator(bool produceToken)
        {
            Location Start = CurrentLocation;
            Token Token = null;

            Debug.Assert(CanStartLineTerminator());

            if (PeekChar() == (char)(65535))
            {
                Token = new EndOfStreamToken(SpanFrom(Start));
            }
            else
            {
                if (ReadChar() == (char)(13))
                {
                    // A CR/LF pair is a legal line terminator
                    if (PeekChar() == (char)(10))
                    {
                        ReadChar();
                    }
                }

                if (produceToken)
                {
                    Token = new LineTerminatorToken(SpanFrom(Start));
                }
                _Line += 1;
                _Column = 1;
            }

            return Token;
        }

        private bool CanStartLineTerminator()
        {
            char c = PeekChar();

            return 
                c == (char)(13) || 
                c == (char)(10) || 
                c == (char)(8232) || 
                c == (char)(8233) || 
                c == (char)(65535);
        }

        private bool EatWhitespace()
        {
            bool functionReturnValue = false;
            char c = PeekChar();

            while (c == (char)(9) || char.GetUnicodeCategory(c) == UnicodeCategory.SpaceSeparator)
            {
                ReadChar();
                functionReturnValue = true;
                c = PeekChar();
            }
            return functionReturnValue;
        }

        private Token Read(bool Advance)
        {
            Token TokenRead;
            Token CurrentToken;

            if (Position > -1)
                CurrentToken = _Tokens[Position];
            else
                CurrentToken = null;

            // If we've reached the end of the stream, just return the end of stream token again
            if (CurrentToken != null && CurrentToken.Type == TokenType.EndOfStream)
                return CurrentToken;

            // If we haven't read a token yet, or if we've reached the end of the tokens that we've read
            // so far, then we need to read a fresh token in.
            if (Position == _Tokens.Count - 1)
            {
            ContinueLine:
                EatWhitespace();

                if (CanStartLineTerminator())
                {
                    TokenRead = ScanLineTerminator();
                }
                else if (CanStartComment())
                {
                    TokenRead = ScanComment();
                }
                else if (CanStartIdentifier())
                {
                    Token Token = ScanIdentifier();

                    if (Token == null)
                        // This was a line continuation, so skip and keep going
                        goto ContinueLine;
                    else
                        TokenRead = Token;
                }
                else if (CanStartStringLiteral())
                {
                    TokenRead = ScanStringLiteral();
                }
                else if (CanStartDateLiteral())
                {
                    TokenRead = ScanDateLiteral();
                }
                else if (CanStartNumericLiteral())
                {
                    TokenRead = ScanNumericLiteral();
                }
                else
                {
                    Location Start = CurrentLocation;
                    TokenType Punctuator = PunctuatorToken.Parse(PeekChar().ToString());

                    if (Punctuator != TokenType.None)
                    {
                        // CONSIDER: Only call this if we know it can start a two-character punctuator
                        TokenRead = ScanPossibleMultiCharacterPunctuator(ReadChar(), Start);
                    }
                    else
                    {
                        ReadChar();
                        TokenRead = new ErrorToken(SyntaxErrorType.InvalidCharacter, SpanFrom(Start));
                    }
                }

                _Tokens.Add(TokenRead);
            }

            // Advance to the next token if we need to
            int NewPosition = Position + 1;
            if (Advance)
                Position = NewPosition;

            return _Tokens[NewPosition];
        }

        #endregion

        #region Shared Methods

        private static bool CanStartNonEscapedIdentifier(char c)
        {
            UnicodeCategory CharClass = char.GetUnicodeCategory(c);
            return IsAlphaClass(CharClass) || IsUnderscoreClass(CharClass);
        }

        private static bool CanStartIdentifier(char c)
        {
            return IsLeftBracket(c) || CanStartNonEscapedIdentifier(c);
        }

        private static bool IsAlphaClass(UnicodeCategory c)
        {
            return c == UnicodeCategory.UppercaseLetter || c == UnicodeCategory.LowercaseLetter || c == UnicodeCategory.TitlecaseLetter || c == UnicodeCategory.OtherLetter || c == UnicodeCategory.ModifierLetter || c == UnicodeCategory.LetterNumber;
        }

        private static bool IsNumericClass(UnicodeCategory c)
        {
            return c == UnicodeCategory.DecimalDigitNumber;
        }

        private static bool IsUnderscoreClass(UnicodeCategory c)
        {
            return c == UnicodeCategory.ConnectorPunctuation;
        }

        private static bool IsSingleQuote(char c)
        {
            return c == '\'' || c == (char)(65287) || c == (char)(8216) || c == (char)(8217);
        }

        private static bool IsDoubleQuote(char c)
        {
            return c == '"' || c == (char)(65282) || c == (char)(8220) || c == (char)(8221);
        }

        private static bool IsDigit(char c)
        {
            return (c >= '0' && c <= '9') || (c >= (char)(65296) && c <= (char)(65305));
        }

        private static bool IsOctalDigit(char c)
        {
            return (c >= '0' && c <= '7') || (c >= (char)(65296) && c <= (char)(65303));
        }

        private static bool IsHexDigit(char c)
        {
            return IsDigit(c) || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F') || (c >= (char)(65345) && c <= (char)(65350)) || (c >= (char)(65313) && c <= (char)(65318));
        }

        private static bool IsEquals(char c)
        {
            return c == '=' || c == (char)(65309);
        }

        private static bool IsLessThan(char c)
        {
            return c == '<' || c == (char)(65308);
        }

        private static bool IsGreaterThan(char c)
        {
            return c == '>' || c == (char)(65310);
        }

        private static bool IsAmpersand(char c)
        {
            return c == '&' || c == (char)(65286);
        }

        private static bool IsUnderscore(char c)
        {
            return IsUnderscoreClass(char.GetUnicodeCategory(c));
        }

        private static bool IsHexDesignator(char c)
        {
            return c == 'H' || c == 'h' || c == (char)(65352) || c == (char)(65320);
        }

        private static bool IsOctalDesignator(char c)
        {
            return c == 'O' || c == 'o' || c == (char)(65327) || c == (char)(65359);
        }

        private static bool IsPeriod(char c)
        {
            return c == '.' || c == (char)(65294);
        }

        private static bool IsExponentDesignator(char c)
        {
            return c == 'e' || c == 'E' || c == (char)(65349) || c == (char)(65317);
        }

        private static bool IsPlus(char c)
        {
            return c == '+' || c == (char)(65291);
        }

        private static bool IsMinus(char c)
        {
            return c == '-' || c == (char)(65293);
        }

        private static bool IsForwardSlash(char c)
        {
            return c == '/' || c == (char)(65295);
        }

        private static bool IsColon(char c)
        {
            return c == ':' || c == (char)(65306);
        }

        private static bool IsPound(char c)
        {
            return c == '#' || c == (char)(65283);
        }

        private static bool IsA(char c)
        {
            return c == 'a' || c == (char)(65345) || c == 'A' || c == (char)(65313);
        }

        private static bool IsP(char c)
        {
            return c == 'p' || c == (char)(65360) || c == 'P' || c == (char)(65328);
        }

        private static bool IsM(char c)
        {
            return c == 'm' || c == (char)(65357) || c == 'M' || c == (char)(65325);
        }

        private static bool IsCharDesignator(char c)
        {
            return c == 'c' || c == 'C' || c == (char)(65347) || c == (char)(65315);
        }

        private static bool IsLeftBracket(char c)
        {
            return c == '[' || c == (char)(65339);
        }

        private static bool IsRightBracket(char c)
        {
            return c == ']' || c == (char)(65341);
        }

        private static bool IsUnsignedTypeChar(char c)
        {
            return c == 'u' || c == 'U' || c == (char)(65333) || c == (char)(65365);
        }

        private static bool IsIdentifier(char c)
        {
            UnicodeCategory CharClass = char.GetUnicodeCategory(c);

            return IsAlphaClass(CharClass) || IsNumericClass(CharClass) || CharClass == UnicodeCategory.SpacingCombiningMark || CharClass == UnicodeCategory.NonSpacingMark || CharClass == UnicodeCategory.Format || IsUnderscoreClass(CharClass);
        }

        private static bool IsTerminatorToken(Token token)
        {
            return
                token.Type == TokenType.EndOfStream ||
                token.Type == TokenType.Colon ||
                token.Type == TokenType.LineTerminator ||
                token.Type == TokenType.Comment;
        }

        static internal char MakeHalfWidth(char c)
        {
            if (c < (char)(65281) || c > (char)(65374))
            {
                return c;
            }
            else
            {
                return (char)((int)(c) - 65280 + 32);
            }
        }

        static internal char MakeFullWidth(char c)
        {
            if (c < (char)(33) || c > (char)(126))
            {
                return c;
            }
            else
            {
                return (char)((int)(c) + 65280 - 32);
            }
        }

        static internal string MakeFullWidth(string s)
        {
            StringBuilder Builder = new StringBuilder(s);

            for (int Index = 0; Index <= Builder.Length - 1; Index++)
            {
                Builder[Index] = MakeFullWidth(Builder[Index]);
            }

            return Builder.ToString();
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Seeks backwards in the stream position to a particular token.
        /// </summary>
        /// <param name="token">The token to seek back to.</param>
        /// <exception cref="System.ObjectDisposedException">Thrown when the scanner has been closed.</exception>
        /// <exception cref="ArgumentException">Thrown when token was not produced by this scanner.</exception>
        public void Seek(Token token)
        {
            int CurrentPosition;
            int StartPosition = Position;
            bool TokenFound = false;

            if (_Disposed) throw new ObjectDisposedException("Scanner");

            if (StartPosition == _Tokens.Count - 1)
                StartPosition -= 1;

            for (CurrentPosition = StartPosition; CurrentPosition >= -1; CurrentPosition += -1)
            {
                if (object.ReferenceEquals(_Tokens[CurrentPosition + 1], token))
                {
                    TokenFound = true;
                    break;
                }
            }

            if (!TokenFound)
                throw new ArgumentException("Token not created by this scanner.");
            else
                Position = CurrentPosition;
        }

        public void Move(Token token)
        {
            int position = -1;
            bool found = false;
            foreach (Token item in _Tokens)
            {
                if (object.ReferenceEquals(item, token))
                {
                    found = true;
                    break;
                }
                position++;
            }

            if (!found)
                throw new ArgumentException("Token not created by this scanner.");
            else
                Position = position;
        }

        /// <summary>
        /// Fetches the previous token in the stream, by using the position entered, ignoring comments, line-breaks and colons
        /// </summary>
        /// <param name="jump">The number of positions that the cursor will be moved back</param>
        /// <param name="move">Move the cursor to the previous position</param>
        /// <returns>The previous token using the specified position.</returns>
        /// <exception cref="System.ObjectDisposedException">Thrown when the scanner has been closed.</exception>
        /// <exception cref="System.ArgumentOutOfRangeException">Thrown when the position entered is out of range.</exception>
        public Token Previous(int jump, bool move)
        {
            if (_Disposed) throw new ObjectDisposedException("Scanner");

            int CurrentPosition = Position + 1;
            int Destination = CurrentPosition - jump;
            Token Current = null;

            if (jump < 1)
            {
                throw new ArgumentOutOfRangeException("jump", "Enter positive positions.");
            }
            else if (Destination >= 0 && Destination < _Tokens.Count)
            {
                if (move)
                    Position = Destination - 1;
                Current = _Tokens[Destination];
            }

            return Current;
        }

        /// <summary>
        /// Fetches the previous token in the stream.
        /// </summary>
        /// <param name="move">Move the cursor to the previous position</param>
        /// <returns>The previous token.</returns>
        /// <exception cref="System.ObjectDisposedException">Thrown when the scanner has been closed.</exception>
        public Token Previous(bool move)
        {
            return Previous(1, move);
        }

        /// <summary>
        /// Fetches the previous token in the stream.
        /// </summary>
        /// <returns>The previous token.</returns>
        /// <exception cref="System.ObjectDisposedException">Thrown when the scanner has been closed.</exception>
        public Token Previous()
        {
            return Previous(false);
        }

        /// <summary>
        /// Try to fetch the previous token in the stream, by using the position entered.
        /// </summary>
        /// <param name="jump">The number of positions that the cursor will be moved back</param>
        /// <param name="move">Move the cursor to the previous position</param>
        /// <returns>The previous token using the specified position, or null if the position is not valid</returns>
        public Token TryPrevious(int jump, bool move)
        {
            try
            {
                return Previous(jump, move);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Fetches the current token without advancing the stream position.
        /// </summary>
        /// <returns>The current token.</returns>
        /// <exception cref="System.ObjectDisposedException">Thrown when the scanner has been closed.</exception>
        public Token Peek()
        {
            if (_Disposed) throw new ObjectDisposedException("Scanner");
            return Read(false);
        }

        /// <summary>
        /// Fetches the current token and advances the stream position.
        /// </summary>
        /// <returns>The current token.</returns>
        /// <exception cref="System.ObjectDisposedException">Thrown when the scanner has been closed.</exception>
        public Token Read()
        {
            if (_Disposed) throw new ObjectDisposedException("Scanner");
            return Read(true);
        }

        /// <summary>
        /// Fetches more than one token at a time from the stream.
        /// </summary>
        /// <param name="buffer">The array to put the tokens into.</param>
        /// <param name="index">The location in the array to start putting the tokens into.</param>
        /// <param name="count">The number of tokens to read.</param>
        /// <returns>The number of tokens read.</returns>
        /// <exception cref="System.ObjectDisposedException">Thrown when the scanner has been closed.</exception>
        /// <exception cref="System.NullReferenceException">Thrown when the buffer is Nothing.</exception>
        /// <exception cref="System.ArgumentException">Thrown when the index or count is invalid, or when there isn't enough room in the buffer.</exception>
        public int ReadBlock(Token[] buffer, int index, int count)
        {
            int FinalCount = 0;

            if (buffer == null) throw new ArgumentNullException("buffer");
            if ((index < 0) || (count < 0)) throw new ArgumentException("Index or count cannot be negative.");
            if (buffer.Length - index < count) throw new ArgumentException("Not enough room in buffer.");
            if (_Disposed) throw new ObjectDisposedException("Scanner");

            while (count > 0)
            {
                Token CurrentToken = Read();

                if (CurrentToken.Type == TokenType.EndOfStream)
                {
                    return FinalCount;
                }

                buffer[FinalCount + index] = CurrentToken;
                count -= 1;
                FinalCount += 1;
            }

            return FinalCount;
        }

        /// <summary>
        /// Reads all of the tokens between the current position and the end of the line (or the end of the stream).
        /// </summary>
        /// <returns>The tokens read.</returns>
        /// <exception cref="System.ObjectDisposedException">Thrown when the scanner has been closed.</exception>
        public Token[] ReadLine()
        {
            List<Token> TokenArray = new List<Token>();

            if (_Disposed) throw new ObjectDisposedException("Scanner");

            while (Peek().Type != TokenType.EndOfStream & Peek().Type != TokenType.LineTerminator)
            {
                TokenArray.Add(Read());
            }

            return TokenArray.ToArray();
        }

        /// <summary>
        /// Reads all the tokens between the current position and the end of the stream.
        /// </summary>
        /// <returns>The tokens read.</returns>
        /// <exception cref="System.ObjectDisposedException">Thrown when the scanner has been closed.</exception>
        public List<Token> ReadToEnd()
        {
            List<Token> TokenArray = new List<Token>();

            if (_Disposed) throw new ObjectDisposedException("Scanner");

            while (Peek().Type != TokenType.EndOfStream)
                TokenArray.Add(Read());

            return TokenArray;
        }

        /// <summary>
        /// Closes/disposes the scanner.
        /// </summary>
        public void Dispose()
        {
            if (!_Disposed)
            {
                _Disposed = true;
                this.Source.Close();
            }
        }
        #endregion

    }
}
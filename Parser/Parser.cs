using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime;
using System.Text;

using Microsoft.VisualBasic.CompilerServices;

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
using VBConverter.CodeParser.Context;
using VBConverter.CodeParser.Error;
using VBConverter.CodeParser.Tokens;
using VBConverter.CodeParser.Tokens.Literal;
using VBConverter.CodeParser.Tokens.Terminator;
using VBConverter.CodeParser.Trees;
using VBConverter.CodeParser.Trees.Arguments;
using VBConverter.CodeParser.Trees.Attributes;
using VBConverter.CodeParser.Trees.CaseClause;
using VBConverter.CodeParser.Trees.Collections;
using VBConverter.CodeParser.Trees.Comments;
using VBConverter.CodeParser.Trees.Declarations;
using VBConverter.CodeParser.Trees.Declarations.Members;
using VBConverter.CodeParser.Trees.Expressions;
using VBConverter.CodeParser.Trees.Imports;
using VBConverter.CodeParser.Trees.Initializers;
using VBConverter.CodeParser.Trees.Modifiers;
using VBConverter.CodeParser.Trees.Names;
using VBConverter.CodeParser.Trees.Parameters;
using VBConverter.CodeParser.Trees.Statements;
using VBConverter.CodeParser.Trees.TypeNames;
using VBConverter.CodeParser.Trees.TypeParameters;
using VBConverter.CodeParser.Trees.Types;
using VBConverter.CodeParser.Trees.VariableDeclarators;

namespace VBConverter.CodeParser
{
    /// <summary>
    /// A parser for the Visual Basic .NET language based on the grammar
    /// documented in the Language Specification.
    /// </summary>
    public sealed class Parser : IDisposable
    {

        #region Fields
        // The scanner we're going to be using.
        private Scanner Scanner;

        // The error table for the parsing
        private IList<SyntaxError> ErrorTable;

        // External line mappings
        private IList<ExternalLineMapping> ExternalLineMappings;

        // Source regions
        private IList<SourceRegion> SourceRegions;

        // External checksums
        private IList<ExternalChecksum> ExternalChecksums;

        // Conditional compilation constants
        private IDictionary<string, object> ConditionalCompilationConstants;

        // Whether there is an error in the construct
        private bool ErrorInConstruct;

        // Whether we're at the beginning of a line
        private bool AtBeginningOfLine;

        // Whether we're doing preprocessing or not
        private bool Preprocess;

        // The current stack of block contexts
        private Stack<TreeType> BlockContextStack = new Stack<TreeType>();

        // The current external source context
        private ExternalSourceContext CurrentExternalSourceContext;

        // The current stack of region contexts
        private Stack<RegionContext> RegionContextStack = new Stack<RegionContext>();

        // The current stack of conditional compilation states
        private Stack<ConditionalCompilationContext> ConditionalCompilationContextStack = new Stack<ConditionalCompilationContext>();

        // Determine whether we have been disposed already or not
        private bool Disposed = false;

        private bool IsSubCall = false;
        private bool HasParenthesisOnSubCall = false;
        private bool HasCallToken = false;

        #endregion

        #region Properties

        private bool ImplementsVB60
        {
            get { return LanguageInfo.ImplementsVB60(Scanner.Version); }
        }

        private bool ImplementsVB71
        {
            get { return LanguageInfo.ImplementsVB71(Scanner.Version); }
        }

        private bool ImplementsVB80
        {
            get { return LanguageInfo.ImplementsVB80(Scanner.Version); }
        }

        #endregion

        #region Dispose Method
        /// <summary>
        /// Disposes the parser.
        /// </summary>
        public void Dispose()
        {
            if (!Disposed)
            {
                Disposed = true;
                Scanner.Dispose();
            }
        }
        #endregion

        #region Token Reading Functions

        private Token Peek()
        {
            return Scanner.Peek();
        }

        private Token PeekAheadOne()
        {
            Token Start = Read();
            Token NextToken = Peek();

            GoBack(Start);
            return NextToken;
        }

        private TokenType PeekAheadFor(params TokenType[] tokens)
        {
            Token Start = Peek();
            Token Current = Start;

            while (!CanEndStatement(Current))
            {
                foreach (TokenType Token in tokens)
                {
                    if (Current.AsUnreservedKeyword() == Token)
                    {
                        GoBack(Start);
                        return Token;
                    }
                }

                Current = Read();
            }

            GoBack(Start);
            return TokenType.None;
        }

        private Token Read()
        {
            return Scanner.Read();
        }

        private Location ReadLocation()
        {
            return Read().Span.Start;
        }

        private Location PeekLocation()
        {
            return Peek().Span.Start;
        }

        private void GoBack(Token token)
        {
            Scanner.Seek(token);
        }

        private void Move(Token token)
        {
            if (token != null)
                Scanner.Move(token);
        }

        private Token MoveToBeginOfStatement()
        {
            Token Current = Peek();
            
            while (true)
            {
                if (Current == null)
                    break;

                if (IsBeginOfStatement(Current))
                    break;

                Current = Previous(true);
            }

            return Current;
        }

        private Token Previous()
        {
            return Scanner.Previous();
        }

        private Token Previous(bool move)
        {
            return Scanner.Previous(move);
        }

        private Token Previous(Token token)
        {
            Token Start = Peek();
            Move(token);
            Token PreviousToken = Scanner.Previous();
            Move(Start);

            return PreviousToken;
        }

        private Token Next(Token token)
        {
            Token Start = Peek();
            Move(token);
            Token NextToken = PeekAheadOne();
            Move(Start);

            return NextToken;
        }

        private void ResyncAt(params TokenType[] tokenTypes)
        {
            Token CurrentToken = Peek();

            while (
                CurrentToken.Type != TokenType.Colon && 
                CurrentToken.Type != TokenType.EndOfStream && 
                CurrentToken.Type != TokenType.LineTerminator && 
                !BeginsStatement(CurrentToken)
                )
            {
                foreach (TokenType Item in tokenTypes)
                {
                    // CONSIDER: Need to check for non-reserved tokens?
                    if (CurrentToken.Type == Item)
                        return;
                }

                Read();
                CurrentToken = Peek();
            }
        }

        private List<Comment> ParseTrailingComments()
        {
            List<Comment> Comments = new List<Comment>();

            // Link in comments that follow the statement
            while (Peek().Type == TokenType.Comment)
            {
                CommentToken token = (CommentToken)Scanner.Read();
                Comments.Add(new Comment(token.Comment, token.IsREM, token.Span));
            }

            if (Comments.Count > 0)
                return Comments;
            
            return null;
        }

        #endregion

        #region Helpers

        private void PushBlockContext(TreeType type)
        {
            BlockContextStack.Push(type);
        }

        private void PopBlockContext()
        {
            BlockContextStack.Pop();
        }

        private TreeType CurrentBlockContextType()
        {
            if (BlockContextStack.Count == 0)
                return TreeType.SyntaxError;
            else
                return BlockContextStack.Peek();
        }

        private static bool IsBinaryOperator(Token token)
        {
            return GetBinaryOperator(token.Type, true) != BinaryOperatorType.None;
        }

        private static bool IsUnaryOperator(Token token)
        {
            return
                token.Type == TokenType.Not ||
                token.Type == TokenType.Minus;
        }

        private static bool IsIdentifier(Token token)
        {
            return
                token != null &&
                IdentifierToken.IsIdentifier(token.Type);
        }

        private static bool IsEvaluable(Token token)
        {
            return
                token != null &&
                IdentifierToken.IsEvaluable(token.Type);
        }

        private static UnaryOperatorType GetUnaryOperator(TokenType type)
        {
            switch (type)
            {
                case TokenType.Plus:
                    return UnaryOperatorType.UnaryPlus;

                case TokenType.Minus:
                    return UnaryOperatorType.Negate;

                case TokenType.Not:
                    return UnaryOperatorType.Not;

                default:
                    return UnaryOperatorType.None;
            }
        }

        private static BinaryOperatorType GetBinaryOperator(TokenType type)
        {
            return GetBinaryOperator(type, false);
        }

        private static BinaryOperatorType GetBinaryOperator(TokenType type, bool allowRange)
        {
            switch (type)
            {
                case TokenType.Ampersand:
                    return BinaryOperatorType.Concatenate;

                case TokenType.Star:
                    return BinaryOperatorType.Multiply;

                case TokenType.ForwardSlash:
                    return BinaryOperatorType.Divide;

                case TokenType.BackwardSlash:
                    return BinaryOperatorType.IntegralDivide;

                case TokenType.Caret:
                    return BinaryOperatorType.Power;

                case TokenType.Plus:
                    return BinaryOperatorType.Plus;

                case TokenType.Minus:
                    return BinaryOperatorType.Minus;

                case TokenType.LessThan:
                    return BinaryOperatorType.LessThan;

                case TokenType.LessThanEquals:
                    return BinaryOperatorType.LessThanEquals;

                case TokenType.Equals:
                    return BinaryOperatorType.Equals;

                case TokenType.NotEquals:
                    return BinaryOperatorType.NotEquals;

                case TokenType.GreaterThan:
                    return BinaryOperatorType.GreaterThan;

                case TokenType.GreaterThanEquals:
                    return BinaryOperatorType.GreaterThanEquals;

                case TokenType.LessThanLessThan:
                    return BinaryOperatorType.ShiftLeft;

                case TokenType.GreaterThanGreaterThan:
                    return BinaryOperatorType.ShiftRight;

                case TokenType.Mod:
                    return BinaryOperatorType.Modulus;

                case TokenType.Or:
                    return BinaryOperatorType.Or;

                case TokenType.OrElse:
                    return BinaryOperatorType.OrElse;

                case TokenType.And:
                    return BinaryOperatorType.And;

                case TokenType.AndAlso:
                    return BinaryOperatorType.AndAlso;

                case TokenType.Xor:
                    return BinaryOperatorType.Xor;

                case TokenType.Like:
                    return BinaryOperatorType.Like;

                case TokenType.Is:
                    return BinaryOperatorType.Is;

                case TokenType.IsNot:
                    return BinaryOperatorType.IsNot;

                case TokenType.To:
                    if (allowRange)
                        return BinaryOperatorType.To;
                    else
                        return BinaryOperatorType.None;
                default:
                    return BinaryOperatorType.None;
            }
        }

        private static PrecedenceLevel GetOperatorPrecedence(UnaryOperatorType @operator)
        {
            switch (@operator)
            {
                case UnaryOperatorType.Not:
                    return PrecedenceLevel.Not;

                case UnaryOperatorType.Negate:
                case UnaryOperatorType.UnaryPlus:
                    return PrecedenceLevel.Negate;

                default:
                    return PrecedenceLevel.None;

            }
        }

        private static PrecedenceLevel GetOperatorPrecedence(BinaryOperatorType @operator)
        {
            switch (@operator)
            {
                case BinaryOperatorType.To:
                    return PrecedenceLevel.Range;

                case BinaryOperatorType.Power:
                    return PrecedenceLevel.Power;

                case BinaryOperatorType.Multiply:
                case BinaryOperatorType.Divide:
                    return PrecedenceLevel.Multiply;

                case BinaryOperatorType.IntegralDivide:
                    return PrecedenceLevel.IntegralDivide;

                case BinaryOperatorType.Modulus:
                    return PrecedenceLevel.Modulus;

                case BinaryOperatorType.Plus:
                case BinaryOperatorType.Minus:
                    return PrecedenceLevel.Plus;

                case BinaryOperatorType.Concatenate:
                    return PrecedenceLevel.Concatenate;

                case BinaryOperatorType.ShiftLeft:
                case BinaryOperatorType.ShiftRight:
                    return PrecedenceLevel.Shift;

                case BinaryOperatorType.Equals:
                case BinaryOperatorType.NotEquals:
                case BinaryOperatorType.LessThan:
                case BinaryOperatorType.GreaterThan:
                case BinaryOperatorType.GreaterThanEquals:
                case BinaryOperatorType.LessThanEquals:
                case BinaryOperatorType.Is:
                case BinaryOperatorType.IsNot:
                case BinaryOperatorType.Like:
                    return PrecedenceLevel.Relational;

                case BinaryOperatorType.And:
                case BinaryOperatorType.AndAlso:
                    return PrecedenceLevel.And;

                case BinaryOperatorType.Or:
                case BinaryOperatorType.OrElse:
                    return PrecedenceLevel.Or;

                case BinaryOperatorType.Xor:
                    return PrecedenceLevel.Xor;

                default:
                    return PrecedenceLevel.None;
            }
        }

        private static BinaryOperatorType GetCompoundAssignmentOperatorType(TokenType tokenType)
        {
            switch (tokenType)
            {
                case TokenType.PlusEquals:
                    return BinaryOperatorType.Plus;

                case TokenType.AmpersandEquals:
                    return BinaryOperatorType.Concatenate;

                case TokenType.StarEquals:
                    return BinaryOperatorType.Multiply;

                case TokenType.MinusEquals:
                    return BinaryOperatorType.Minus;

                case TokenType.ForwardSlashEquals:
                    return BinaryOperatorType.Divide;

                case TokenType.BackwardSlashEquals:
                    return BinaryOperatorType.IntegralDivide;

                case TokenType.CaretEquals:
                    return BinaryOperatorType.Power;

                case TokenType.LessThanLessThanEquals:
                    return BinaryOperatorType.ShiftLeft;

                case TokenType.GreaterThanGreaterThanEquals:
                    return BinaryOperatorType.ShiftRight;

                default:
                    return BinaryOperatorType.None;
            }
        }

        private static TreeType GetAssignmentOperator(TokenType tokenType)
        {

            switch (tokenType)
            {
                case TokenType.Equals:
                    return TreeType.AssignmentStatement;

                case TokenType.PlusEquals:
                case TokenType.AmpersandEquals:
                case TokenType.StarEquals:
                case TokenType.MinusEquals:
                case TokenType.ForwardSlashEquals:
                case TokenType.BackwardSlashEquals:
                case TokenType.CaretEquals:
                case TokenType.LessThanLessThanEquals:
                case TokenType.GreaterThanGreaterThanEquals:
                    return TreeType.CompoundAssignmentStatement;

                default:
                    break;
            }

            return TreeType.SyntaxError;
        }

        private static bool IsRelationalOperator(TokenType type)
        {
            return type >= TokenType.LessThan && type <= TokenType.GreaterThanEquals;
        }

        private static bool IsOverloadableOperator(Token op)
        {
            switch (op.Type)
            {
                case TokenType.Plus:
                case TokenType.Minus:
                case TokenType.Not:
                case TokenType.Star:
                case TokenType.ForwardSlash:
                case TokenType.BackwardSlash:
                case TokenType.Ampersand:
                case TokenType.Like:
                case TokenType.Mod:
                case TokenType.And:
                case TokenType.Or:
                case TokenType.Xor:
                case TokenType.Caret:
                case TokenType.LessThanLessThan:
                case TokenType.GreaterThanGreaterThan:
                case TokenType.Equals:
                case TokenType.NotEquals:
                case TokenType.LessThan:
                case TokenType.GreaterThan:
                case TokenType.LessThanEquals:
                case TokenType.GreaterThanEquals:
                case TokenType.CType:
                    return true;
                case TokenType.Identifier:
                    if (op.AsUnreservedKeyword() == TokenType.IsTrue || op.AsUnreservedKeyword() == TokenType.IsFalse)
                        return true;

                    break;
            }

            return false;
        }

        private static BlockType GetContinueType(TokenType tokenType)
        {
            switch (tokenType)
            {
                case TokenType.Do:
                    return BlockType.Do;

                case TokenType.For:
                    return BlockType.For;

                case TokenType.While:
                    return BlockType.While;

                default:
                    return BlockType.None;
            }
        }

        private static BlockType GetExitType(TokenType tokenType)
        {
            switch (tokenType)
            {
                case TokenType.Do:
                    return BlockType.Do;

                case TokenType.For:
                    return BlockType.For;

                case TokenType.While:
                    return BlockType.While;

                case TokenType.Select:
                    return BlockType.Select;

                case TokenType.Sub:
                    return BlockType.Sub;

                case TokenType.Function:
                    return BlockType.Function;

                case TokenType.Property:
                    return BlockType.Property;

                case TokenType.Try:
                    return BlockType.Try;

                default:
                    return BlockType.None;
            }
        }

        private static BlockType GetBlockType(TokenType type)
        {
            switch (type)
            {
                case TokenType.Wend:
                case TokenType.While:
                    return BlockType.While;

                case TokenType.Select:
                    return BlockType.Select;

                case TokenType.If:
                    return BlockType.If;

                case TokenType.Try:
                    return BlockType.Try;

                case TokenType.SyncLock:
                    return BlockType.SyncLock;

                case TokenType.Using:
                    return BlockType.Using;

                case TokenType.With:
                    return BlockType.With;

                case TokenType.Sub:
                    return BlockType.Sub;

                case TokenType.Function:
                    return BlockType.Function;

                case TokenType.Operator:
                    return BlockType.Operator;

                case TokenType.Get:
                    return BlockType.Get;

                case TokenType.Set:
                    return BlockType.Set;

                case TokenType.Event:
                    return BlockType.Event;

                case TokenType.AddHandler:
                    return BlockType.AddHandler;

                case TokenType.RemoveHandler:
                    return BlockType.RemoveHandler;

                case TokenType.RaiseEvent:
                    return BlockType.RaiseEvent;

                case TokenType.Property:
                    return BlockType.Property;

                case TokenType.Class:
                    return BlockType.Class;

                case TokenType.Structure:
                case TokenType.Type:
                    return BlockType.Structure;

                case TokenType.Module:
                    return BlockType.Module;

                case TokenType.Interface:
                    return BlockType.Interface;

                case TokenType.Enum:
                    return BlockType.Enum;

                case TokenType.Namespace:
                    return BlockType.Namespace;

                default:
                    return BlockType.None;
            }
        }

        private void ReportSyntaxError(SyntaxError syntaxError)
        {
            if (ErrorInConstruct)
                return;

            ErrorInConstruct = true;

            if (ErrorTable != null)
                ErrorTable.Add(syntaxError);
        }

        private void ReportSyntaxError(SyntaxErrorType errorType, Span span)
        {
            ReportSyntaxError(new SyntaxError(errorType, span));
        }

        private void ReportSyntaxError(SyntaxErrorType errorType, Token firstToken, Token lastToken)
        {
            // A lexical error takes precedence over the parser error
            if (firstToken.Type == TokenType.LexicalError)
                ReportSyntaxError(((ErrorToken)firstToken).SyntaxError);
            else
                ReportSyntaxError(errorType, SpanFrom(firstToken, lastToken));
        }

        private void ReportSyntaxError(SyntaxErrorType errorType, Token token)
        {
            ReportSyntaxError(errorType, token, token);
        }

        private static bool StatementEndsBlock(TreeType blockStatementType, Statement endStatement)
        {
            switch (blockStatementType)
            {
                case TreeType.DoBlockStatement:
                    return endStatement.Type == TreeType.LoopStatement;

                case TreeType.ForBlockStatement:
                case TreeType.ForEachBlockStatement:
                    return endStatement.Type == TreeType.NextStatement;

                case TreeType.WhileBlockStatement:
                    return (endStatement.Type == TreeType.EndBlockStatement) && ((EndBlockStatement)endStatement).EndType == BlockType.While;

                case TreeType.SyncLockBlockStatement:
                    return (endStatement.Type == TreeType.EndBlockStatement) && ((EndBlockStatement)endStatement).EndType == BlockType.SyncLock;

                case TreeType.UsingBlockStatement:
                    return (endStatement.Type == TreeType.EndBlockStatement) && ((EndBlockStatement)endStatement).EndType == BlockType.Using;

                case TreeType.WithBlockStatement:
                    return (endStatement.Type == TreeType.EndBlockStatement) && ((EndBlockStatement)endStatement).EndType == BlockType.With;

                case TreeType.TryBlockStatement:
                case TreeType.CatchBlockStatement:
                    return endStatement.Type == TreeType.CatchStatement || endStatement.Type == TreeType.FinallyStatement || (endStatement.Type == TreeType.EndBlockStatement && ((EndBlockStatement)endStatement).EndType == BlockType.Try);

                case TreeType.FinallyBlockStatement:
                    return endStatement.Type == TreeType.EndBlockStatement && ((EndBlockStatement)endStatement).EndType == BlockType.Try;

                case TreeType.SelectBlockStatement:
                case TreeType.CaseBlockStatement:
                    return endStatement.Type == TreeType.CaseStatement || endStatement.Type == TreeType.CaseElseStatement || (endStatement.Type == TreeType.EndBlockStatement && ((EndBlockStatement)endStatement).EndType == BlockType.Select);

                case TreeType.CaseElseBlockStatement:
                    return endStatement.Type == TreeType.EndBlockStatement && ((EndBlockStatement)endStatement).EndType == BlockType.Select;

                case TreeType.IfBlockStatement:
                case TreeType.ElseIfBlockStatement:
                    return endStatement.Type == TreeType.ElseIfStatement || endStatement.Type == TreeType.ElseStatement || (endStatement.Type == TreeType.EndBlockStatement && ((EndBlockStatement)endStatement).EndType == BlockType.If);

                case TreeType.ElseBlockStatement:
                    return endStatement.Type == TreeType.EndBlockStatement && ((EndBlockStatement)endStatement).EndType == BlockType.If;

                case TreeType.LineIfBlockStatement:
                    return false;

                case TreeType.SubDeclaration:
                case TreeType.ConstructorDeclaration:
                    return endStatement.Type == TreeType.EndBlockStatement && ((EndBlockStatement)endStatement).EndType == BlockType.Sub;

                case TreeType.FunctionDeclaration:
                    return endStatement.Type == TreeType.EndBlockStatement && ((EndBlockStatement)endStatement).EndType == BlockType.Function;

                case TreeType.OperatorDeclaration:
                    return endStatement.Type == TreeType.EndBlockStatement && ((EndBlockStatement)endStatement).EndType == BlockType.Operator;

                case TreeType.GetAccessorDeclaration:
                    return endStatement.Type == TreeType.EndBlockStatement && ((EndBlockStatement)endStatement).EndType == BlockType.Get;

                case TreeType.SetAccessorDeclaration:
                    return endStatement.Type == TreeType.EndBlockStatement && ((EndBlockStatement)endStatement).EndType == BlockType.Set;

                case TreeType.PropertyDeclaration:
                    return endStatement.Type == TreeType.EndBlockStatement && ((EndBlockStatement)endStatement).EndType == BlockType.Property;

                case TreeType.CustomEventDeclaration:
                    return endStatement.Type == TreeType.EndBlockStatement && ((EndBlockStatement)endStatement).EndType == BlockType.Event;

                case TreeType.AddHandlerAccessorDeclaration:
                    return endStatement.Type == TreeType.EndBlockStatement && ((EndBlockStatement)endStatement).EndType == BlockType.AddHandler;

                case TreeType.RemoveHandlerAccessorDeclaration:
                    return endStatement.Type == TreeType.EndBlockStatement && ((EndBlockStatement)endStatement).EndType == BlockType.RemoveHandler;

                case TreeType.RaiseEventAccessorDeclaration:
                    return endStatement.Type == TreeType.EndBlockStatement && ((EndBlockStatement)endStatement).EndType == BlockType.RaiseEvent;

                case TreeType.ClassDeclaration:
                    return endStatement.Type == TreeType.EndBlockStatement && ((EndBlockStatement)endStatement).EndType == BlockType.Class;

                case TreeType.StructureDeclaration:
                    return endStatement.Type == TreeType.EndBlockStatement && ((EndBlockStatement)endStatement).EndType == BlockType.Structure;

                case TreeType.ModuleDeclaration:
                    return endStatement.Type == TreeType.EndBlockStatement && ((EndBlockStatement)endStatement).EndType == BlockType.Module;

                case TreeType.InterfaceDeclaration:
                    return endStatement.Type == TreeType.EndBlockStatement && ((EndBlockStatement)endStatement).EndType == BlockType.Interface;

                case TreeType.EnumDeclaration:
                    return endStatement.Type == TreeType.EndBlockStatement && ((EndBlockStatement)endStatement).EndType == BlockType.Enum;

                case TreeType.NamespaceDeclaration:
                    return endStatement.Type == TreeType.EndBlockStatement && ((EndBlockStatement)endStatement).EndType == BlockType.Namespace;

                default:
                    Debug.Fail("Unexpected.");
                    break;
            }

            return false;
        }

        private static bool DeclarationEndsBlock(TreeType blockDeclarationType, EndBlockDeclaration endDeclaration)
        {
            switch (blockDeclarationType)
            {
                case TreeType.SubDeclaration:
                case TreeType.ConstructorDeclaration:
                    return endDeclaration.EndType == BlockType.Sub;

                case TreeType.FunctionDeclaration:
                    return endDeclaration.EndType == BlockType.Function;

                case TreeType.OperatorDeclaration:
                    return endDeclaration.EndType == BlockType.Operator;

                case TreeType.PropertyDeclaration:
                    return endDeclaration.EndType == BlockType.Property;

                case TreeType.GetAccessorDeclaration:
                    return endDeclaration.EndType == BlockType.Get;

                case TreeType.SetAccessorDeclaration:
                    return endDeclaration.EndType == BlockType.Set;

                case TreeType.CustomEventDeclaration:
                    return endDeclaration.EndType == BlockType.Event;

                case TreeType.AddHandlerAccessorDeclaration:
                    return endDeclaration.EndType == BlockType.AddHandler;

                case TreeType.RemoveHandlerAccessorDeclaration:
                    return endDeclaration.EndType == BlockType.RemoveHandler;

                case TreeType.RaiseEventAccessorDeclaration:
                    return endDeclaration.EndType == BlockType.RaiseEvent;

                case TreeType.ClassDeclaration:
                    return endDeclaration.EndType == BlockType.Class;

                case TreeType.StructureDeclaration:
                    return endDeclaration.EndType == BlockType.Structure;

                case TreeType.ModuleDeclaration:
                    return endDeclaration.EndType == BlockType.Module;

                case TreeType.InterfaceDeclaration:
                    return endDeclaration.EndType == BlockType.Interface;

                case TreeType.EnumDeclaration:
                    return endDeclaration.EndType == BlockType.Enum;

                case TreeType.NamespaceDeclaration:
                    return endDeclaration.EndType == BlockType.Namespace;

                default:
                    Debug.Fail("Unexpected.");
                    break;
            }

            return false;
        }

        private bool ValidInContext(TreeType blockType, TreeType declarationType)
        {
            TreeType[] AllowedMembers = null;
            
            switch (blockType)
            {
                case TreeType.FileTree:
                    if (ImplementsVB60)
                    {
                        AllowedMembers = new TreeType[] 
                        { 
                            TreeType.EmptyDeclaration,
                            TreeType.OptionDeclaration,
                            TreeType.StructureDeclaration,
                            TreeType.EnumDeclaration,
                            TreeType.SubDeclaration,
                            TreeType.FunctionDeclaration,
                            TreeType.ConstructorDeclaration,
                            TreeType.PropertyDeclaration,
                            TreeType.VariableListDeclaration,
                            TreeType.ExternalSubDeclaration,
                            TreeType.ExternalFunctionDeclaration,
                            TreeType.ImplementsDeclaration,
                            TreeType.EventDeclaration
                        };
                    }
                    else
                    {
                        AllowedMembers = new TreeType[] 
                        { 
                            TreeType.EmptyDeclaration,
                            TreeType.OptionDeclaration,
                            TreeType.StructureDeclaration,
                            TreeType.EnumDeclaration,
                            TreeType.ImportsDeclaration,
                            TreeType.AttributeDeclaration,
                            TreeType.NamespaceDeclaration,
                            TreeType.ClassDeclaration,
                            TreeType.InterfaceDeclaration,
                            TreeType.DelegateSubDeclaration,
                            TreeType.DelegateFunctionDeclaration,
                            TreeType.ModuleDeclaration
                        };
                    }

                    break;
                case TreeType.NamespaceDeclaration:
                    AllowedMembers = new TreeType[] 
                        { 
                            TreeType.NamespaceDeclaration,
                            TreeType.ClassDeclaration,
                            TreeType.StructureDeclaration,
                            TreeType.InterfaceDeclaration,
                            TreeType.DelegateSubDeclaration,
                            TreeType.DelegateFunctionDeclaration,
                            TreeType.EnumDeclaration,
                            TreeType.ModuleDeclaration
                        };

                    break;
                case TreeType.ClassDeclaration:
                    AllowedMembers = new TreeType[] 
                        { 
                            TreeType.ClassDeclaration,
                            TreeType.StructureDeclaration,
                            TreeType.InterfaceDeclaration,
                            TreeType.DelegateSubDeclaration,
                            TreeType.DelegateFunctionDeclaration,
                            TreeType.EnumDeclaration,
                            TreeType.EventDeclaration,
                            TreeType.SubDeclaration,
                            TreeType.FunctionDeclaration,
                            TreeType.PropertyDeclaration,
                            TreeType.CustomEventDeclaration,
                            TreeType.VariableListDeclaration,
                            TreeType.ExternalSubDeclaration,
                            TreeType.ExternalFunctionDeclaration,
                            TreeType.ConstructorDeclaration,
                            TreeType.ImplementsDeclaration,
                            TreeType.OperatorDeclaration,
                            TreeType.InheritsDeclaration
                        };

                    break;
                case TreeType.StructureDeclaration:
                    AllowedMembers = new TreeType[] 
                        { 
                            TreeType.ClassDeclaration,
                            TreeType.StructureDeclaration,
                            TreeType.InterfaceDeclaration,
                            TreeType.DelegateSubDeclaration,
                            TreeType.DelegateFunctionDeclaration,
                            TreeType.EnumDeclaration,
                            TreeType.EventDeclaration,
                            TreeType.SubDeclaration,
                            TreeType.FunctionDeclaration,
                            TreeType.PropertyDeclaration,
                            TreeType.CustomEventDeclaration,
                            TreeType.VariableListDeclaration,
                            TreeType.ExternalSubDeclaration,
                            TreeType.ExternalFunctionDeclaration,
                            TreeType.ConstructorDeclaration,
                            TreeType.ImplementsDeclaration,
                            TreeType.OperatorDeclaration
                        };

                    // Type Block declarations (VB6) only allow variable declarations inside
                    if (ImplementsVB60)
                        AllowedMembers = new TreeType[] { TreeType.VariableListDeclaration };

                    break;
                case TreeType.ModuleDeclaration:
                    AllowedMembers = new TreeType[] 
                        { 
                            TreeType.ClassDeclaration,
                            TreeType.StructureDeclaration,
                            TreeType.InterfaceDeclaration,
                            TreeType.DelegateSubDeclaration,
                            TreeType.DelegateFunctionDeclaration,
                            TreeType.EnumDeclaration,
                            TreeType.EventDeclaration,
                            TreeType.SubDeclaration,
                            TreeType.FunctionDeclaration,
                            TreeType.PropertyDeclaration,
                            TreeType.CustomEventDeclaration,
                            TreeType.VariableListDeclaration,
                            TreeType.ExternalSubDeclaration,
                            TreeType.ExternalFunctionDeclaration
                        };

                    break;
                case TreeType.InterfaceDeclaration:
                    AllowedMembers = new TreeType[] 
                        { 
                            TreeType.ClassDeclaration,
                            TreeType.StructureDeclaration,
                            TreeType.InterfaceDeclaration,
                            TreeType.DelegateSubDeclaration,
                            TreeType.DelegateFunctionDeclaration,
                            TreeType.EnumDeclaration,
                            TreeType.EventDeclaration,
                            TreeType.SubDeclaration,
                            TreeType.FunctionDeclaration,
                            TreeType.PropertyDeclaration,
                            TreeType.InheritsDeclaration
                        };

                    break;
                case TreeType.CustomEventDeclaration:
                    AllowedMembers = new TreeType[] 
                        { 
                            TreeType.AddHandlerAccessorDeclaration,
                            TreeType.RemoveHandlerAccessorDeclaration,
                            TreeType.RaiseEventAccessorDeclaration
                        };

                    break;
                case TreeType.PropertyDeclaration:
                    AllowedMembers = new TreeType[] 
                        { 
                            TreeType.GetAccessorDeclaration,
                            TreeType.SetAccessorDeclaration
                        };

                    break;
                case TreeType.EnumDeclaration:
                    AllowedMembers = new TreeType[] 
                        { 
                            TreeType.EnumValueDeclaration
                        };
                    break;
                case TreeType.EmptyDeclaration:
                    AllowedMembers = new TreeType[] 
                        { 
                            TreeType.EmptyDeclaration
                        };
                    break;
            }

            bool Result = false;
            
            if (declarationType == TreeType.EmptyDeclaration)
                Result = true;
            else if (AllowedMembers != null)
                Result = Array.Exists<TreeType>(AllowedMembers, delegate(TreeType i) { return i == declarationType; });

            return Result;
        }

        private SyntaxErrorType ValidDeclaration(TreeType blockType, Declaration declaration, List<Declaration> declarations)
        {
            SyntaxErrorType Result = SyntaxErrorType.None;
            bool IsValid = ValidInContext(blockType, declaration.Type);
            
            if (!IsValid)
                return InvalidDeclarationTypeError(blockType);

            SyntaxErrorType errorOnFail = SyntaxErrorType.None;
            TreeType[] expectedTypes = null;
            TreeType[] unexpectedTypes = null;

            switch (declaration.Type)
            {
                case TreeType.InheritsDeclaration:
                    if (blockType != TreeType.ClassDeclaration)
                    {
                        Result = SyntaxErrorType.InheritsMustBeFirst;
                    }
                    else if (blockType != TreeType.InterfaceDeclaration)
                    {
                        InheritsDeclaration Current = (InheritsDeclaration)declaration;
                        if (Current.InheritedTypes.Count > 1)
                            Result = SyntaxErrorType.NoMultipleInheritance;
                    }
                    else
                    {
                        errorOnFail = SyntaxErrorType.InheritsMustBeFirst;
                        expectedTypes = new TreeType[] { TreeType.EmptyDeclaration, TreeType.InheritsDeclaration };
                    }

                    break;
                case TreeType.ImplementsDeclaration:
                    if (ImplementsVB60)
                    {
                        errorOnFail = SyntaxErrorType.InvalidDeclarationInFile;
                        unexpectedTypes = new TreeType[] { TreeType.PropertyDeclaration, TreeType.SubDeclaration, TreeType.FunctionDeclaration, TreeType.ConstructorDeclaration };
                    }
                    else
                    {
                        errorOnFail = SyntaxErrorType.ImplementsInWrongOrder;
                        expectedTypes = new TreeType[] { TreeType.EmptyDeclaration, TreeType.InheritsDeclaration, TreeType.ImplementsDeclaration };
                    }
                    break;                
                case TreeType.OptionDeclaration:
                    errorOnFail = SyntaxErrorType.OptionStatementWrongOrder;
                    expectedTypes = new TreeType[] { TreeType.EmptyDeclaration, TreeType.OptionDeclaration };
                    break;
                case TreeType.ImportsDeclaration:
                    errorOnFail = SyntaxErrorType.ImportsStatementWrongOrder;
                    expectedTypes = new TreeType[] { TreeType.EmptyDeclaration, TreeType.OptionDeclaration, TreeType.ImportsDeclaration };
                    break;
                case TreeType.AttributeDeclaration:
                    errorOnFail= SyntaxErrorType.ImportsStatementWrongOrder;
                    expectedTypes = new TreeType[] { TreeType.EmptyDeclaration, TreeType.OptionDeclaration, TreeType.ImportsDeclaration, TreeType.AttributeDeclaration };
                    break;
                case TreeType.PropertyDeclaration:
                case TreeType.SubDeclaration:
                case TreeType.FunctionDeclaration:
                case TreeType.EmptyDeclaration:
                case TreeType.ConstructorDeclaration:
                    errorOnFail = SyntaxErrorType.None;
                    break;
                default:
                    if (ImplementsVB60)
                    {
                        errorOnFail = SyntaxErrorType.InvalidDeclarationInFile;
                        unexpectedTypes = new TreeType[] { TreeType.PropertyDeclaration, TreeType.SubDeclaration, TreeType.FunctionDeclaration, TreeType.Comment };
                    }
                    else
                    {
                        errorOnFail = SyntaxErrorType.None;
                    }
                    break;
            }

            if (Result == SyntaxErrorType.None && errorOnFail != SyntaxErrorType.None)
                Result = ValidateTree(errorOnFail, declarations, expectedTypes, unexpectedTypes);

            return Result;
        }

        private static SyntaxErrorType ValidateTree(SyntaxErrorType errorOnFail, ICollection declarations, TreeType[] expectedTypes, TreeType[] unexpectedTypes)
        {
            if (declarations.Count == 0 || (expectedTypes == null && unexpectedTypes == null))
                return SyntaxErrorType.None;

            SyntaxErrorType result = SyntaxErrorType.None;

            foreach (Declaration Item in declarations)
            {
                if (expectedTypes != null)
                {
                    result = errorOnFail;
                    foreach (TreeType ValidType in expectedTypes)
                    {
                        if (Item.Type == ValidType)
                        {
                            result = SyntaxErrorType.None;
                            break;
                        }
                    }
                }

                if (result == SyntaxErrorType.None && unexpectedTypes != null)
                {
                    result = SyntaxErrorType.None;
                    foreach (TreeType InvalidType in unexpectedTypes)
                    {
                        if (Item.Type == InvalidType)
                        {
                            result = errorOnFail;
                            break;
                        }
                    }
                }
            }

            return result;
        }

        private void ReportMismatchedEndError(TreeType blockType, Span actualEndSpan)
        {
            SyntaxErrorType ErrorType = SyntaxErrorType.None;

            switch (blockType)
            {
                case TreeType.DoBlockStatement:
                    ErrorType = SyntaxErrorType.ExpectedLoop;
                    break;

                case TreeType.ForBlockStatement:
                case TreeType.ForEachBlockStatement:
                    ErrorType = SyntaxErrorType.ExpectedNext;
                    break;

                case TreeType.WhileBlockStatement:
                    ErrorType = SyntaxErrorType.ExpectedEndWhile;
                    break;

                case TreeType.SelectBlockStatement:
                case TreeType.CaseBlockStatement:
                case TreeType.CaseElseBlockStatement:
                    ErrorType = SyntaxErrorType.ExpectedEndSelect;
                    break;

                case TreeType.SyncLockBlockStatement:
                    ErrorType = SyntaxErrorType.ExpectedEndSyncLock;
                    break;

                case TreeType.UsingBlockStatement:
                    ErrorType = SyntaxErrorType.ExpectedEndUsing;
                    break;

                case TreeType.IfBlockStatement:
                case TreeType.ElseIfBlockStatement:
                case TreeType.ElseBlockStatement:
                    ErrorType = SyntaxErrorType.ExpectedEndIf;
                    break;

                case TreeType.TryBlockStatement:
                case TreeType.CatchBlockStatement:
                case TreeType.FinallyBlockStatement:
                    ErrorType = SyntaxErrorType.ExpectedEndTry;
                    break;

                case TreeType.WithBlockStatement:
                    ErrorType = SyntaxErrorType.ExpectedEndWith;
                    break;

                case TreeType.SubDeclaration:
                case TreeType.ConstructorDeclaration:
                    ErrorType = SyntaxErrorType.ExpectedEndSub;
                    break;

                case TreeType.FunctionDeclaration:
                    ErrorType = SyntaxErrorType.ExpectedEndFunction;
                    break;

                case TreeType.OperatorDeclaration:
                    ErrorType = SyntaxErrorType.ExpectedEndOperator;
                    break;

                case TreeType.PropertyDeclaration:
                    ErrorType = SyntaxErrorType.ExpectedEndProperty;
                    break;

                case TreeType.GetAccessorDeclaration:
                    ErrorType = SyntaxErrorType.ExpectedEndGet;
                    break;

                case TreeType.SetAccessorDeclaration:
                    ErrorType = SyntaxErrorType.ExpectedEndSet;
                    break;

                case TreeType.CustomEventDeclaration:
                    ErrorType = SyntaxErrorType.ExpectedEndEvent;
                    break;

                case TreeType.AddHandlerAccessorDeclaration:
                    ErrorType = SyntaxErrorType.ExpectedEndAddHandler;
                    break;

                case TreeType.RemoveHandlerAccessorDeclaration:
                    ErrorType = SyntaxErrorType.ExpectedEndRemoveHandler;
                    break;

                case TreeType.RaiseEventAccessorDeclaration:
                    ErrorType = SyntaxErrorType.ExpectedEndRaiseEvent;
                    break;

                case TreeType.ClassDeclaration:
                    ErrorType = SyntaxErrorType.ExpectedEndClass;
                    break;

                case TreeType.StructureDeclaration:
                    ErrorType = SyntaxErrorType.ExpectedEndStructure;
                    break;

                case TreeType.ModuleDeclaration:
                    ErrorType = SyntaxErrorType.ExpectedEndModule;
                    break;

                case TreeType.InterfaceDeclaration:
                    ErrorType = SyntaxErrorType.ExpectedEndInterface;
                    break;

                case TreeType.EnumDeclaration:
                    ErrorType = SyntaxErrorType.ExpectedEndEnum;
                    break;

                case TreeType.NamespaceDeclaration:
                    ErrorType = SyntaxErrorType.ExpectedEndNamespace;
                    break;

                default:
                    Debug.Fail("Unexpected.");
                    break;
            }

            ReportSyntaxError(ErrorType, actualEndSpan);
        }

        private void ReportMissingBeginStatementError(TreeType blockStatementType, Statement endStatement)
        {
            SyntaxErrorType ErrorType = SyntaxErrorType.None;

            switch (endStatement.Type)
            {
                case TreeType.LoopStatement:
                    ErrorType = SyntaxErrorType.LoopWithoutDo;
                    break;

                case TreeType.NextStatement:
                    ErrorType = SyntaxErrorType.NextWithoutFor;
                    break;

                case TreeType.EndBlockStatement:
                    switch (((EndBlockStatement)endStatement).EndType)
                    {
                        case BlockType.While:
                            ErrorType = SyntaxErrorType.EndWhileWithoutWhile;
                            break;

                        case BlockType.Select:
                            ErrorType = SyntaxErrorType.EndSelectWithoutSelect;
                            break;

                        case BlockType.SyncLock:
                            ErrorType = SyntaxErrorType.EndSyncLockWithoutSyncLock;
                            break;

                        case BlockType.Using:
                            ErrorType = SyntaxErrorType.EndUsingWithoutUsing;
                            break;

                        case BlockType.If:
                            ErrorType = SyntaxErrorType.EndIfWithoutIf;
                            break;

                        case BlockType.Try:
                            ErrorType = SyntaxErrorType.EndTryWithoutTry;
                            break;

                        case BlockType.With:
                            ErrorType = SyntaxErrorType.EndWithWithoutWith;
                            break;

                        case BlockType.Sub:
                            ErrorType = SyntaxErrorType.EndSubWithoutSub;
                            break;

                        case BlockType.Function:
                            ErrorType = SyntaxErrorType.EndFunctionWithoutFunction;
                            break;

                        case BlockType.Operator:
                            ErrorType = SyntaxErrorType.EndOperatorWithoutOperator;
                            break;

                        case BlockType.Get:
                            ErrorType = SyntaxErrorType.EndGetWithoutGet;
                            break;

                        case BlockType.Set:
                            ErrorType = SyntaxErrorType.EndSetWithoutSet;
                            break;

                        case BlockType.Property:
                            ErrorType = SyntaxErrorType.EndPropertyWithoutProperty;
                            break;

                        case BlockType.Event:
                            ErrorType = SyntaxErrorType.EndEventWithoutEvent;
                            break;

                        case BlockType.AddHandler:
                            ErrorType = SyntaxErrorType.EndAddHandlerWithoutAddHandler;
                            break;

                        case BlockType.RemoveHandler:
                            ErrorType = SyntaxErrorType.EndRemoveHandlerWithoutRemoveHandler;
                            break;

                        case BlockType.RaiseEvent:
                            ErrorType = SyntaxErrorType.EndRaiseEventWithoutRaiseEvent;
                            break;

                        case BlockType.Class:
                            ErrorType = SyntaxErrorType.EndClassWithoutClass;
                            break;

                        case BlockType.Structure:
                            ErrorType = SyntaxErrorType.EndStructureWithoutStructure;
                            break;

                        case BlockType.Module:
                            ErrorType = SyntaxErrorType.EndModuleWithoutModule;
                            break;

                        case BlockType.Interface:
                            ErrorType = SyntaxErrorType.EndInterfaceWithoutInterface;
                            break;

                        case BlockType.Enum:
                            ErrorType = SyntaxErrorType.EndEnumWithoutEnum;
                            break;

                        case BlockType.Namespace:
                            ErrorType = SyntaxErrorType.EndNamespaceWithoutNamespace;
                            break;

                        default:
                            Debug.Fail("Unexpected.");
                            break;
                    }
                    break;

                case TreeType.CatchStatement:
                    if (blockStatementType == TreeType.FinallyBlockStatement)
                    {
                        ErrorType = SyntaxErrorType.CatchAfterFinally;
                    }
                    else
                    {
                        ErrorType = SyntaxErrorType.CatchWithoutTry;
                    }

                    break;

                case TreeType.FinallyStatement:
                    if (blockStatementType == TreeType.FinallyBlockStatement)
                    {
                        ErrorType = SyntaxErrorType.FinallyAfterFinally;
                    }
                    else
                    {
                        ErrorType = SyntaxErrorType.FinallyWithoutTry;
                    }

                    break;

                case TreeType.CaseStatement:
                    if (blockStatementType == TreeType.CaseElseBlockStatement)
                    {
                        ErrorType = SyntaxErrorType.CaseAfterCaseElse;
                    }
                    else
                    {
                        ErrorType = SyntaxErrorType.CaseWithoutSelect;
                    }

                    break;

                case TreeType.CaseElseStatement:
                    if (blockStatementType == TreeType.CaseElseBlockStatement)
                    {
                        ErrorType = SyntaxErrorType.CaseElseAfterCaseElse;
                    }
                    else
                    {
                        ErrorType = SyntaxErrorType.CaseElseWithoutSelect;
                    }

                    break;

                case TreeType.ElseIfStatement:
                    if (blockStatementType == TreeType.ElseBlockStatement)
                    {
                        ErrorType = SyntaxErrorType.ElseIfAfterElse;
                    }
                    else
                    {
                        ErrorType = SyntaxErrorType.ElseIfWithoutIf;
                    }

                    break;

                case TreeType.ElseStatement:
                    if (blockStatementType == TreeType.ElseBlockStatement)
                    {
                        ErrorType = SyntaxErrorType.ElseAfterElse;
                    }
                    else
                    {
                        ErrorType = SyntaxErrorType.ElseWithoutIf;
                    }

                    break;

                default:
                    Debug.Fail("Unexpected.");
                    break;
            }

            ReportSyntaxError(ErrorType, endStatement.Span);
        }

        private void ReportMissingBeginDeclarationError(EndBlockDeclaration endDeclaration)
        {
            SyntaxErrorType ErrorType = SyntaxErrorType.None;

            switch (endDeclaration.EndType)
            {
                case BlockType.Sub:
                    ErrorType = SyntaxErrorType.EndSubWithoutSub;
                    break;

                case BlockType.Function:
                    ErrorType = SyntaxErrorType.EndFunctionWithoutFunction;
                    break;

                case BlockType.Operator:
                    ErrorType = SyntaxErrorType.EndOperatorWithoutOperator;
                    break;

                case BlockType.Property:
                    ErrorType = SyntaxErrorType.EndPropertyWithoutProperty;
                    break;

                case BlockType.Get:
                    ErrorType = SyntaxErrorType.EndGetWithoutGet;
                    break;

                case BlockType.Set:
                    ErrorType = SyntaxErrorType.EndSetWithoutSet;
                    break;

                case BlockType.Event:
                    ErrorType = SyntaxErrorType.EndEventWithoutEvent;
                    break;

                case BlockType.AddHandler:
                    ErrorType = SyntaxErrorType.EndAddHandlerWithoutAddHandler;
                    break;

                case BlockType.RemoveHandler:
                    ErrorType = SyntaxErrorType.EndRemoveHandlerWithoutRemoveHandler;
                    break;

                case BlockType.RaiseEvent:
                    ErrorType = SyntaxErrorType.EndRaiseEventWithoutRaiseEvent;
                    break;

                case BlockType.Class:
                    ErrorType = SyntaxErrorType.EndClassWithoutClass;
                    break;

                case BlockType.Structure:
                    ErrorType = SyntaxErrorType.EndStructureWithoutStructure;
                    break;

                case BlockType.Module:
                    ErrorType = SyntaxErrorType.EndModuleWithoutModule;
                    break;

                case BlockType.Interface:
                    ErrorType = SyntaxErrorType.EndInterfaceWithoutInterface;
                    break;

                case BlockType.Enum:
                    ErrorType = SyntaxErrorType.EndEnumWithoutEnum;
                    break;

                case BlockType.Namespace:
                    ErrorType = SyntaxErrorType.EndNamespaceWithoutNamespace;
                    break;

                default:
                    Debug.Fail("Unexpected.");
                    break;
            }

            ReportSyntaxError(ErrorType, endDeclaration.Span);
        }

        private static SyntaxErrorType InvalidDeclarationTypeError(TreeType blockType)
        {
            SyntaxErrorType ErrorType = SyntaxErrorType.None;

            switch (blockType)
            {
                case TreeType.PropertyDeclaration:
                    ErrorType = SyntaxErrorType.InvalidInsideProperty;
                    break;

                case TreeType.ClassDeclaration:
                    ErrorType = SyntaxErrorType.InvalidInsideClass;
                    break;

                case TreeType.StructureDeclaration:
                    ErrorType = SyntaxErrorType.InvalidInsideStructure;
                    break;

                case TreeType.ModuleDeclaration:
                    ErrorType = SyntaxErrorType.InvalidInsideModule;
                    break;

                case TreeType.InterfaceDeclaration:
                    ErrorType = SyntaxErrorType.InvalidInsideInterface;
                    break;

                case TreeType.EnumDeclaration:
                    ErrorType = SyntaxErrorType.InvalidInsideEnum;
                    break;

                case TreeType.NamespaceDeclaration:
                case TreeType.FileTree:
                    ErrorType = SyntaxErrorType.InvalidInsideNamespace;
                    break;

                default:
                    Debug.Fail("Unexpected.");
                    break;
            }

            return ErrorType;
        }

        private void HandleUnexpectedToken(TokenType type)
        {
            SyntaxErrorType ErrorType = SyntaxErrorType.None;

            switch (type)
            {
                case TokenType.Comma:
                    ErrorType = SyntaxErrorType.ExpectedComma;
                    break;

                case TokenType.LeftParenthesis:
                    ErrorType = SyntaxErrorType.ExpectedLeftParenthesis;
                    break;

                case TokenType.RightParenthesis:
                    ErrorType = SyntaxErrorType.ExpectedRightParenthesis;
                    break;

                case TokenType.Equals:
                    ErrorType = SyntaxErrorType.ExpectedEquals;
                    break;

                case TokenType.As:
                    ErrorType = SyntaxErrorType.ExpectedAs;
                    break;

                case TokenType.RightCurlyBrace:
                    ErrorType = SyntaxErrorType.ExpectedRightCurlyBrace;
                    break;

                case TokenType.Period:
                    ErrorType = SyntaxErrorType.ExpectedPeriod;
                    break;

                case TokenType.Minus:
                    ErrorType = SyntaxErrorType.ExpectedMinus;
                    break;

                case TokenType.Is:
                    ErrorType = SyntaxErrorType.ExpectedIs;
                    break;

                case TokenType.GreaterThan:
                    ErrorType = SyntaxErrorType.ExpectedGreaterThan;
                    break;

                case TokenType.Of:
                    ErrorType = SyntaxErrorType.ExpectedOf;
                    break;
                case TokenType.Get:
                case TokenType.Let:
                case TokenType.Set:
                    ErrorType = SyntaxErrorType.ExpectedPropertyAcessor;
                    break;
                default:
                    Debug.Fail("Should give a more specific error.");
                    ErrorType = SyntaxErrorType.SyntaxError;
                    break;
            }

            ReportSyntaxError(ErrorType, Peek());
        }

        private Location VerifyExpectedToken(TokenType type)
        {
            Token Current = Peek();
            if (Current.Type == type)
                return ReadLocation();

            HandleUnexpectedToken(type);

            return Location.Empty;
        }

        private bool VerifyExpectedToken(bool checkUnreserved, params TokenType[] types)
        {
            if (types == null)
                return false;

            Token current = Peek();
            bool exists = false;

            foreach (TokenType type in types)
            {
                if (type == current.Type)
                {
                    exists = true;
                    break;
                }

                if (checkUnreserved)
                {
                    if (type == current.AsUnreservedKeyword())
                    {
                        exists = true;
                        break;
                    }
                }
            }

            if (!exists && types.Length > 0)
                HandleUnexpectedToken(types[0]);

            return exists;
        }

        private bool VerifyExpectedToken(params TokenType[] types)
        {
            return VerifyExpectedToken(false, types);
        }

        private bool CanEndStatement(Token token)
        {
            return
                token == null ||
                IdentifierToken.IsTerminator(token.Type) ||
                (BlockContextStack.Count > 0 && CurrentBlockContextType() == TreeType.LineIfBlockStatement && token.Type == TokenType.Else);
        }

        private bool BeginsStatement(Token token)
        {
            if (!CanEndStatement(token))
                return false;
            
            switch (token.Type)
            {
                case TokenType.AddHandler:
                case TokenType.Call:
                case TokenType.Case:
                case TokenType.Catch:
                case TokenType.Class:
                case TokenType.Const:
                case TokenType.Declare:
                case TokenType.Delegate:
                case TokenType.Dim:
                case TokenType.Do:
                case TokenType.Else:
                case TokenType.ElseIf:
                case TokenType.End:
                case TokenType.EndIf:
                case TokenType.Enum:
                case TokenType.Erase:
                case TokenType.Error:
                case TokenType.Event:
                case TokenType.Exit:
                case TokenType.Finally:
                case TokenType.For:
                case TokenType.Friend:
                case TokenType.Function:
                case TokenType.Get:
                case TokenType.GoTo:
                case TokenType.GoSub:
                case TokenType.If:
                case TokenType.Implements:
                case TokenType.Imports:
                case TokenType.Inherits:
                case TokenType.Interface:
                case TokenType.Loop:
                case TokenType.Module:
                case TokenType.MustInherit:
                case TokenType.MustOverride:
                case TokenType.Namespace:
                case TokenType.Narrowing:
                case TokenType.Next:
                case TokenType.NotInheritable:
                case TokenType.NotOverridable:
                case TokenType.Option:
                case TokenType.Overloads:
                case TokenType.Overridable:
                case TokenType.Overrides:
                case TokenType.Partial:
                case TokenType.Private:
                case TokenType.Property:
                case TokenType.Protected:
                case TokenType.Public:
                case TokenType.RaiseEvent:
                case TokenType.ReadOnly:
                case TokenType.ReDim:
                case TokenType.RemoveHandler:
                case TokenType.Resume:
                case TokenType.Return:
                case TokenType.Select:
                case TokenType.Shadows:
                case TokenType.Shared:
                case TokenType.Static:
                case TokenType.Stop:
                case TokenType.Structure:
                case TokenType.Sub:
                case TokenType.SyncLock:
                case TokenType.Throw:
                case TokenType.Try:
                case TokenType.Type:
                case TokenType.Using:
                case TokenType.Wend:
                case TokenType.While:
                case TokenType.Widening:
                case TokenType.With:
                case TokenType.WithEvents:
                case TokenType.WriteOnly:
                case TokenType.Pound:
                    return true;

                default:
                    return false;
            }
        }

        private void VerifyEndOfCallStatement()
        {
            VerifyEndOfCallStatement(null);
        }

        private void VerifyEndOfCallStatement(CallStatement statement)
        {
            Token Current = Peek();

            if (ImplementsVB60 && !HasCallToken && (!IsSubCall || HasParenthesisOnSubCall))
            {
                Token PreviousToken = Previous(Current);
                bool AllowParenthesis = false;

                if (statement != null && statement.Arguments != null)
                {
                    if (statement.Arguments.Count == 1)
                        AllowParenthesis = true;
                }
                
                if (!AllowParenthesis && PreviousToken.Type == TokenType.RightParenthesis)
                    ReportSyntaxError(SyntaxErrorType.ExpectedEquals, PreviousToken);
            }
        }

        private Token VerifyEndOfStatement()
        {
            return VerifyEndOfStatement(null);
        }

        private Token VerifyEndOfStatement(Statement statement)
        {
            Token Current = Peek();

            Debug.Assert(!(Current.Type == TokenType.Comment), "Should have dealt with these by now!");

            if (statement is CallStatement)
                VerifyEndOfCallStatement((CallStatement)statement);

            if (Current.Type == TokenType.LineTerminator || Current.Type == TokenType.EndOfStream)
            {
                AtBeginningOfLine = true;
            }
            else if (Current.Type == TokenType.Colon || (ImplementsVB60 && statement is LabelStatement))
            {
                AtBeginningOfLine = false;
            }
            else if (Current.Type == TokenType.Else && CurrentBlockContextType() == TreeType.LineIfBlockStatement)
            {
                // Line If statements allow Else to end the statement
                AtBeginningOfLine = false;
            }
            else
            {
                // Syntax error -- a valid statement is followed by something other than
                // a colon, end-of-line, or a comment.
                ResyncAt();
                ReportSyntaxError(SyntaxErrorType.ExpectedEndOfStatement, Current);
                return VerifyEndOfStatement();
            }

            Current = Read();

            HasCallToken = false;
            IsSubCall = false;
            HasParenthesisOnSubCall = false;

            return Current;
        }

        private bool IsFirstTokenOnStatement(Token token)
        {
            Token Start = Peek();
            Move(token);
            Token PreviousToken = Previous();
            Move(Start);

            return
                PreviousToken == null ||
                PreviousToken.Type == TokenType.Colon ||
                PreviousToken.Type == TokenType.LineTerminator ||
                PreviousToken.Type == TokenType.Comment;
        }

        private bool IsBeginOfStatement(Token token)
        {
            //Token Start = Peek();
            //Move(token);
            //Token PreviousToken = Previous();
            //Move(Start);

            bool result = false;

            Token FirstPrevious = Previous(token);
            Token SecondPrevious = Previous(FirstPrevious);

            if (FirstPrevious == null)
                result = true;
            else if (IsEndOfStatement(FirstPrevious))
                result = true;
            else if (ImplementsVB60 && FirstPrevious.Type == TokenType.IntegerLiteral)
                result = IsEndOfStatement(SecondPrevious);

            return result;
        }

        private bool IsEndOfStatement(Token token)
        {
            // one-line staatement: If Condition then DoSomething
            bool IsLineIfStatement = CurrentBlockContextType() == TreeType.LineIfBlockStatement; 

            return
                token == null ||
                IdentifierToken.IsTerminator(token.Type) ||
                (IsLineIfStatement && token.Type == TokenType.Then);
        }

        private Token GetFirstTokenOfStatement(Token token)
        {
            Token Start = Peek();
            Move(token);
            Token Current = MoveToBeginOfStatement();
            Move(Start);

            return Current;
        }

        private Token GetLastTokenOfStatement(Token token)
        {
            Token Start = Peek();
            Move(token);

            Token Current = null;
            Token NextToken = null;
            bool IsEnd = false;

            do
            {
                Current = Read();
                NextToken = Next(Current);
                IsEnd = IsEndOfStatement(NextToken);
            }
            while (!IsEnd);

            Move(Start);

            return Current;
        }

        private Span SpanFrom(Location location)
        {
            Location EndLocation;

            if (Peek().Type == TokenType.EndOfStream)
                EndLocation = Scanner.Previous().Span.Finish;
            else
                EndLocation = Peek().Span.Start;

            return new Span(location, EndLocation);
        }

        private Span SpanFrom(Token token)
        {
            Location StartLocation;
            Location EndLocation;

            if (token.Type == TokenType.EndOfStream && !Scanner.IsOnFirstToken)
            {
                StartLocation = Scanner.Previous().Span.Finish;
                EndLocation = StartLocation;
            }
            else
            {
                StartLocation = token.Span.Start;

                if (Peek().Type == TokenType.EndOfStream && !Scanner.IsOnFirstToken)
                    EndLocation = Scanner.Previous().Span.Finish;
                else
                    EndLocation = Peek().Span.Start;
            }

            return new Span(StartLocation, EndLocation);
        }

        private Span SpanFrom(Token startToken, Token endToken)
        {
            Location StartLocation;
            Location EndLocation;

            if (startToken.Type == TokenType.EndOfStream && !Scanner.IsOnFirstToken)
            {
                StartLocation = Scanner.Previous().Span.Finish;
                EndLocation = StartLocation;
            }
            else
            {
                StartLocation = startToken.Span.Start;

                if (endToken.Type == TokenType.EndOfStream && !Scanner.IsOnFirstToken)
                    EndLocation = Scanner.Previous().Span.Finish;
                else if (endToken.Span.Start.Index == startToken.Span.Start.Index)
                    EndLocation = endToken.Span.Finish;
                else
                    EndLocation = endToken.Span.Start;
            }

            return new Span(StartLocation, EndLocation);
        }

        private static Span SpanFrom(Token startToken, Tree endTree)
        {
            return new Span(startToken.Span.Start, endTree.Span.Finish);
        }

        private Span SpanFrom(Statement startStatement, Statement endStatement)
        {
            if (endStatement == null)
                return SpanFrom(startStatement.Span.Start);
            else
                return new Span(startStatement.Span.Start, endStatement.Span.Start);
        }

        private void NormalizeDeclarationList(List<Declaration> declarations)
        {
            if (!ImplementsVB60)
                return;

            List<int> removeList = new List<int>();
            
            foreach (Declaration firstItem in declarations)
            {
                foreach (Declaration secondItem in declarations)
                {
                    if (!object.ReferenceEquals(firstItem, secondItem) && (firstItem is PropertyDeclaration) && (secondItem is PropertyDeclaration))
                    {
                        PropertyDeclaration firstProperty = (PropertyDeclaration)firstItem;
                        PropertyDeclaration secondProperty = (PropertyDeclaration)secondItem;

                        if (firstProperty.Name == secondProperty.Name && firstProperty.GetAccessor != null)
                        {
                            removeList.Add(secondProperty.GetHashCode());
                            firstProperty.SetAccessor = secondProperty.SetAccessor;
                            break;
                        }
                    }
                }
            }


            while (removeList.Count > 0)
            {
                int removeIndex = -1;
                int currentIndex = -1;
                
                foreach (Declaration item in declarations)
                {
                    currentIndex++;
                    if (removeList[0] == item.GetHashCode())
                    {
                        removeIndex = currentIndex;
                        removeList.RemoveAt(0);
                        break;
                    }
                }

                if (removeIndex >= 0)
                    declarations.RemoveAt(removeIndex);
            }
        }

        #endregion

        #region Names

        private SimpleName ParseSimpleName(bool allowKeyword)
        {
            Token Current = Peek();
            if (Current.Type == TokenType.Identifier)
            {
                IdentifierToken IdentifierToken = (IdentifierToken)Read();
                return new SimpleName(IdentifierToken.Identifier, IdentifierToken.TypeCharacter, IdentifierToken.Escaped, IdentifierToken.Span);
            }
            else
            {
                // If the token is a keyword, assume that the user meant it to 
                // be an identifer and consume it. Otherwise, leave current token
                // as is and let caller decide what to do.
                if (IdentifierToken.IsKeyword(Peek().Type))
                {
                    IdentifierToken Item = (IdentifierToken)Read();
                    if (!allowKeyword)
                        ReportSyntaxError(SyntaxErrorType.InvalidUseOfKeyword, Item);

                    return new SimpleName(Item.Identifier, Item.TypeCharacter, Item.Escaped, Item.Span);
                }
                else
                {
                    ReportSyntaxError(SyntaxErrorType.ExpectedIdentifier, Current);
                    return SimpleName.GetBadSimpleName(SpanFrom(Peek(), Current));
                }
            }
        }

        private Name ParseName(bool AllowGlobal)
        {
            Token Start = Peek();
            Name Result;
            bool QualificationRequired = false;

            // Seeing the global token implies > LanguageVersion.VisualBasic71
            if (Start.Type == TokenType.Global)
            {
                if (!AllowGlobal)
                    ReportSyntaxError(SyntaxErrorType.InvalidUseOfGlobal, Peek());
                
                Read();
                Result = new SpecialName(TreeType.GlobalNamespaceName, SpanFrom(Start));
                QualificationRequired = true;
            }
            else
            {
                Result = ParseSimpleName(false);
            }

            if (Peek().Type == TokenType.Period)
            {
                SimpleName Qualifier;
                Location DotLocation;

                do
                {
                    DotLocation = ReadLocation();
                    Qualifier = ParseSimpleName(true);
                    Result = new QualifiedName(Result, DotLocation, Qualifier, SpanFrom(Start));
                }
                while (Peek().Type == TokenType.Period);
            }
            else if (QualificationRequired)
            {
                ReportSyntaxError(SyntaxErrorType.ExpectedPeriod, Peek());
            }

            return Result;
        }

        private VariableName ParseVariableName(bool allowExplicitArraySizes)
        {
            Token Start = Peek();
            SimpleName Name;
            ArrayTypeName ArrayType = null;

            // CONSIDER: Often, programmers put extra decl specifiers where they are not required. 
            // Eg: Dim x as Integer, Dim y as Long
            // Check for this and give a more informative error?

            Name = ParseSimpleName(false);

            if (Peek().Type == TokenType.LeftParenthesis)
                ArrayType = ParseArrayTypeName(null, null, allowExplicitArraySizes, false);

            return new VariableName(Name, ArrayType, SpanFrom(Start));
        }

        private Name ParseNameListName()
        {
            return ParseNameListName(false);
        }

        // This function implements some of the special name parsing logic for names in a name
        // list such as Implements and Handles
        private Name ParseNameListName(bool allowLeadingMeOrMyBase)
        {
            Token Start = Peek();
            Name Result;

            if (Start.Type == TokenType.MyBase && allowLeadingMeOrMyBase)
                Result = new SpecialName(TreeType.MyBaseName, SpanFrom(ReadLocation()));
            else if (Start.Type == TokenType.Me && allowLeadingMeOrMyBase && ImplementsVB80)
                Result = new SpecialName(TreeType.MeName, SpanFrom(ReadLocation()));
            else
                Result = ParseSimpleName(false);

            if (Peek().Type == TokenType.Period)
            {
                SimpleName Qualifier;
                Location DotLocation;

                do
                {
                    DotLocation = ReadLocation();
                    Qualifier = ParseSimpleName(true);
                    Result = new QualifiedName(Result, DotLocation, Qualifier, SpanFrom(Start));
                }
                while (Peek().Type == TokenType.Period);
            }
            else
            {
                ReportSyntaxError(SyntaxErrorType.ExpectedPeriod, Peek());
            }

            return Result;
        }

        #endregion

        #region Types

        private TypeName ParseTypeName(bool allowArrayType)
        {
            return ParseTypeName(allowArrayType, false);
        }

        private TypeName ParseTypeName(bool allowArrayType, bool allowOpenType)
        {
            Token Start = Peek();
            TypeName Result = null;
            IntrinsicType Types = IdentifierToken.GetIntrinsicType(Start.Type);

            if (Types == IntrinsicType.None)
            {
                TokenType UnreservedType = Start.AsUnreservedKeyword();
                IntrinsicType UnreservedIntrinsic = IdentifierToken.GetIntrinsicType(UnreservedType);

                if (UnreservedIntrinsic != IntrinsicType.None)
                {
                    Types = UnreservedIntrinsic;
                }
                else if (Start.Type == TokenType.Identifier || Start.Type == TokenType.Global)
                {
                    Result = ParseNamedTypeName(true, allowOpenType);
                }
                else
                {
                    ReportSyntaxError(SyntaxErrorType.ExpectedType, Start);
                    Result = NamedTypeName.GetBadNamedType(SpanFrom(Start));
                }
            }
            
            if (Result == null)
            {
                Read();
                Result = new IntrinsicTypeName(Types, Start.Span);

                // VB6: Fixed length string
                if (ImplementsVB60 && Types == IntrinsicType.String && Peek().Type == TokenType.Star)
                {
                    Read();
                    Expression argument = ParseExpression();
                    Result = new IntrinsicTypeName(IntrinsicType.FixedString, Result.Span);
                    ((IntrinsicTypeName)Result).StringLength = argument;
                }
            }

            if (allowArrayType && Peek().Type == TokenType.LeftParenthesis)
                Result = ParseArrayTypeName(Start, Result, false, false);
            
            return Result;
        }

        private NamedTypeName ParseNamedTypeName(bool allowGlobal)
        {
            return ParseNamedTypeName(allowGlobal, false);
        }

        private NamedTypeName ParseNamedTypeName(bool allowGlobal, bool allowOpenType)
        {
            Token Start = Peek();
            Name Name = ParseName(allowGlobal);

            if (Peek().Type == TokenType.LeftParenthesis)
            {
                Token LeftParenthesis = Read();

                if (Peek().Type == TokenType.Of)
                    return new ConstructedTypeName(Name, ParseTypeArguments(LeftParenthesis, allowOpenType), SpanFrom(Start));
        
                GoBack(LeftParenthesis);
            }

            return new NamedTypeName(Name, Name.Span);
        }

        private TypeArgumentCollection ParseTypeArguments(Token leftParenthesis)
        {
            return ParseTypeArguments(leftParenthesis, false);
        }

        private TypeArgumentCollection ParseTypeArguments(Token leftParenthesis, bool allowOpenType)
        {
            Token Start = leftParenthesis;
            Location OfLocation;
            List<TypeName> TypeArguments = new List<TypeName>();
            List<Location> CommaLocations = new List<Location>();
            Location RightParenthesisLocation;
            bool OpenType = false;

            Debug.Assert(Peek().Type == TokenType.Of);
            OfLocation = ReadLocation();

            if ((Peek().Type == TokenType.Comma || Peek().Type == TokenType.RightParenthesis) && allowOpenType)
            {
                OpenType = true;
            }

            if (!OpenType || Peek().Type != TokenType.RightParenthesis)
            {
                do
                {
                    TypeName TypeArgument;

                    if (TypeArguments.Count > 0 || OpenType)
                        CommaLocations.Add(ReadLocation());
                    
                    if (!OpenType)
                    {
                        TypeArgument = ParseTypeName(true);

                        if (ErrorInConstruct)
                            ResyncAt(TokenType.Comma, TokenType.RightParenthesis);
                        
                        TypeArguments.Add(TypeArgument);
                    }
                }
                while (ArrayBoundsContinue());
            }

            RightParenthesisLocation = VerifyExpectedToken(TokenType.RightParenthesis);
            return new TypeArgumentCollection(OfLocation, TypeArguments, CommaLocations, RightParenthesisLocation, SpanFrom(Start));
        }

        private bool ArrayBoundsContinue()
        {
            Token NextToken = Peek();

            if (NextToken.Type == TokenType.Comma)
            {
                return true;
            }
            else if (NextToken.Type == TokenType.RightParenthesis || IsEndOfStatement(NextToken))
            {
                return false;
            }

            ReportSyntaxError(SyntaxErrorType.ArgumentSyntax, NextToken);
            ResyncAt(TokenType.Comma, TokenType.RightParenthesis);

            if (Peek().Type == TokenType.Comma)
            {
                ErrorInConstruct = false;
                return true;
            }

            return false;
        }

        private ArrayTypeName ParseArrayTypeName(Token startType, TypeName elementType, bool allowExplicitSizes, bool innerArrayType)
        {
            Token ArgumentsStart = Peek();
            List<Argument> Arguments;
            List<Location> CommaLocations = new List<Location>();
            Location RightParenthesisLocation;
            ArgumentCollection ArgumentCollection;

            Debug.Assert(Peek().Type == TokenType.LeftParenthesis);

            if (startType == null)
                startType = Peek();

            Read();

            if (Peek().Type == TokenType.RightParenthesis || Peek().Type == TokenType.Comma)
            {
                Arguments = null;

                while (Peek().Type == TokenType.Comma)
                    CommaLocations.Add(ReadLocation());
            }
            else
            {
                if (!allowExplicitSizes)
                {
                    if (innerArrayType)
                        ReportSyntaxError(SyntaxErrorType.NoConstituentArraySizes, Peek());
                    else
                        ReportSyntaxError(SyntaxErrorType.NoExplicitArraySizes, Peek());
                }

                Arguments = ParseArrayArguments(CommaLocations);
            }

            RightParenthesisLocation = VerifyExpectedToken(TokenType.RightParenthesis);
            ArgumentCollection = new ArgumentCollection(Arguments, CommaLocations, RightParenthesisLocation, SpanFrom(ArgumentsStart));

            if (Peek().Type == TokenType.LeftParenthesis)
            {
                if (ImplementsVB60)
                    ReportSyntaxError(SyntaxErrorType.ExpectedEndOfStatement, Peek().Span);
                else
                    elementType = ParseArrayTypeName(Peek(), elementType, false, true);
            }

            return new ArrayTypeName(elementType, CommaLocations.Count + 1, ArgumentCollection, SpanFrom(startType));
        }

        private List<Argument> ParseArrayArguments(List<Location> commaLocations)
        {
            List<Argument> arguments = new List<Argument>();
            
            do
            {
                if (arguments.Count > 0 && commaLocations != null)
                    commaLocations.Add(ReadLocation());

                bool allowRange = ImplementsVB60 || ImplementsVB80;
                Token start = Peek();
                Expression size = ParseExpression(allowRange);

                if (ErrorInConstruct)
                    ResyncAt(TokenType.Comma, TokenType.RightParenthesis, TokenType.As);

                arguments.Add(new Argument(null, Location.Empty, size, SpanFrom(start)));
            }
            while (ArrayBoundsContinue());
            
            return arguments;
        }

        #endregion

        #region Initializers

        private Initializer ParseInitializer()
        {
            if (Peek().Type == TokenType.LeftCurlyBrace)
            {
                return ParseAggregateInitializer();
            }
            else
            {
                Expression Expression = ParseExpression();
                return new ExpressionInitializer(Expression, Expression.Span);
            }
        }

        private bool InitializersContinue()
        {
            Token NextToken = Peek();

            if (NextToken.Type == TokenType.Comma)
            {
                return true;
            }
            else if (NextToken.Type == TokenType.RightCurlyBrace || IsEndOfStatement(NextToken))
            {
                return false;
            }

            ReportSyntaxError(SyntaxErrorType.InitializerSyntax, NextToken);
            ResyncAt(TokenType.Comma, TokenType.RightCurlyBrace);

            if (Peek().Type == TokenType.Comma)
            {
                ErrorInConstruct = false;
                return true;
            }

            return false;
        }

        private AggregateInitializer ParseAggregateInitializer()
        {
            Token Start = Peek();
            List<Initializer> Initializers = new List<Initializer>();
            Location RightBracketLocation;
            List<Location> CommaLocations = new List<Location>();

            Debug.Assert(Start.Type == TokenType.LeftCurlyBrace);
            Read();

            if (Peek().Type != TokenType.RightCurlyBrace)
            {
                do
                {
                    if (Initializers.Count > 0)
                    {
                        CommaLocations.Add(ReadLocation());
                    }

                    Initializers.Add(ParseInitializer());

                    if (ErrorInConstruct)
                    {
                        ResyncAt(TokenType.Comma, TokenType.RightCurlyBrace);
                    }
                }
                while (InitializersContinue());
            }

            RightBracketLocation = VerifyExpectedToken(TokenType.RightCurlyBrace);

            return new AggregateInitializer(new InitializerCollection(Initializers, CommaLocations, RightBracketLocation, SpanFrom(Start)), SpanFrom(Start));
        }

        #endregion

        #region Arguments

        private Argument ParseArgument(ref bool foundNamedArgument)
        {
            Token Start = Read();
            Expression Value;
            SimpleName Name;
            Location ColonEqualsLocation = Location.Empty;

            if (Peek().Type == TokenType.ColonEquals)
            {
                if (!foundNamedArgument)
                    foundNamedArgument = true;
            }
            else if (foundNamedArgument)
            {
                ReportSyntaxError(SyntaxErrorType.ExpectedNamedArgument, Start);
                foundNamedArgument = false;
            }

            GoBack(Start);
            Name = null;

            Token Current = Peek();
            Token ByValToken = null;

            if (ImplementsVB60 && Current.Type == TokenType.ByVal)
                ByValToken = Read();
            
            if (foundNamedArgument)
            {
                Name = ParseSimpleName(true);
                ColonEqualsLocation = ReadLocation();
            }

            Value = ParseExpression();

            if (ErrorInConstruct)
            {
                ResyncAt(TokenType.Comma, TokenType.RightParenthesis);

                if (Peek().Type == TokenType.Comma)
                    ErrorInConstruct = false;
            }

            Span startSpan = SpanFrom(Start);
            Location byValLocation = Location.Empty;

            if (ImplementsVB60 && ByValToken != null)
                byValLocation = ByValToken.Span.Start;

            return new Argument(Name, ColonEqualsLocation, Value, startSpan, byValLocation);
        }

        private bool ArgumentsContinue()
        {
            Token NextToken = Peek();

            if (NextToken.Type == TokenType.Comma)
            {
                return true;
            }
            else if (NextToken.Type == TokenType.RightParenthesis || IsEndOfStatement(NextToken))
            {
                return false;
            }

            ReportSyntaxError(SyntaxErrorType.ArgumentSyntax, NextToken);
            ResyncAt(TokenType.Comma, TokenType.RightParenthesis);

            if (Peek().Type == TokenType.Comma)
            {
                ErrorInConstruct = false;
                return true;
            }

            return false;
        }

        private bool IsBeginOfArgumentList(Token token, bool bypassParenthesis)
        {
            bool Result = false;
            Token PreviousToken = Previous();
            bool IsPreviousIdentifier = IsIdentifier(PreviousToken);
            
            if (bypassParenthesis)
                Result = IsPreviousIdentifier && ((IdentifierToken)PreviousToken).HasSpaceAfter;
            else
                Result = (token.Type == TokenType.LeftParenthesis);
            
            return Result;
        }

        private bool IsEndOfArgumentList(Token token, bool bypassParenthesis)
        {
            bool Result = false;

            if (bypassParenthesis && IsEndOfStatement(token))
                Result = true; 
            else if (!bypassParenthesis && token.Type == TokenType.RightParenthesis)
                Result = true;
            
            return Result;
        }

        private bool HasCallOnStatement()
        {
            Token Start = Peek();
            Token Current = MoveToBeginOfStatement();
            Move(Start);

            return (Current != null && Current.Type == TokenType.Call);
        }

        private bool UseParenthesisOnStatement(Token Identifier)
        {
            Token NextToken = Next(Identifier);
            return
                (NextToken != null && NextToken.Type == TokenType.LeftParenthesis);
        }

        private bool IsDotBangToken(Token token)
        {
            return
                token != null &&
                (token.Type == TokenType.Period ||
                token.Type == TokenType.Exclamation);
        }

        private ArgumentCollection ParseArguments()
        {
            Token Start = Peek();
            Token Method = Previous();

            bool UseParenthesis = UseParenthesisOnStatement(Method);
            bool IsBegin = IsBeginOfArgumentList(Start, !UseParenthesis);
            
            List<Argument> Arguments = new List<Argument>();
            List<Location> CommaLocations = new List<Location>();
            Location FinishLocation = Location.Empty;

            if (!IsBegin)
                return null;
            else if (UseParenthesis)
                Read();

            Token Last = Peek();
            bool IsEnd = IsEndOfArgumentList(Last, !UseParenthesis);

            if (!IsEnd)
            {
                bool FoundNamedArgument = false;
                Token ArgumentStart = null;
                Argument Argument = null;

                do
                {
                    if (Arguments.Count > 0)
                        CommaLocations.Add(ReadLocation());

                    ArgumentStart = Peek();

                    if (ArgumentStart.Type == TokenType.Comma || ArgumentStart.Type == TokenType.RightParenthesis)
                    {
                        if (FoundNamedArgument)
                            ReportSyntaxError(SyntaxErrorType.ExpectedNamedArgument, Peek());

                        Argument = null;
                    }
                    else
                    {
                        Argument = ParseArgument(ref FoundNamedArgument);
                    }

                    Arguments.Add(Argument);
                }
                while (ArgumentsContinue());
            }

            if (IsEndOfArgumentList(Peek(), !UseParenthesis))
            {
                if (UseParenthesis)
                    Read();

                FinishLocation = PeekLocation();
            }
            else 
            {
                Token Current = Peek();

                // On error, peek for ")" with "(". If ")" seen before 
                // "(", then sync on that. Otherwise, assume missing ")"
                // and let caller decide.
                ResyncAt(TokenType.LeftParenthesis, TokenType.RightParenthesis);

                if (Peek().Type == TokenType.RightParenthesis)
                {
                    ReportSyntaxError(SyntaxErrorType.SyntaxError, Peek());
                    FinishLocation = ReadLocation();
                }
                else
                {
                    GoBack(Current);
                    ReportSyntaxError(SyntaxErrorType.ExpectedRightParenthesis, Peek());
                }
            }

            return new ArgumentCollection(Arguments, CommaLocations, FinishLocation, SpanFrom(Start));
        }

        #endregion

        #region Expressions

        private Expression ParseLiteralExpression()
        {
            Token Start = Read();

            switch (Start.Type)
            {
                case TokenType.True:
                case TokenType.False:
                    return new BooleanLiteralExpression(Start.Type == TokenType.True, Start.Span);

                case TokenType.IntegerLiteral:
                    IntegerLiteralToken Literal = (IntegerLiteralToken)Start;
                    return new IntegerLiteralExpression(Literal.Literal, Literal.IntegerBase, Literal.TypeCharacter, Literal.Span);

                case TokenType.FloatingPointLiteral:
                    FloatingPointLiteralToken FloatingPointLiteral = (FloatingPointLiteralToken)Start;
                    return new FloatingPointLiteralExpression(FloatingPointLiteral.Literal, FloatingPointLiteral.TypeCharacter, FloatingPointLiteral.Span);

                case TokenType.DecimalLiteral:
                    DecimalLiteralToken DecimalLiteral = (DecimalLiteralToken)Start;
                    return new DecimalLiteralExpression(DecimalLiteral.Literal, DecimalLiteral.TypeCharacter, DecimalLiteral.Span);

                case TokenType.CharacterLiteral:
                    CharacterLiteralToken CharacterLiteral = (CharacterLiteralToken)Start;
                    return new CharacterLiteralExpression(CharacterLiteral.Literal, CharacterLiteral.Span);

                case TokenType.StringLiteral:
                    StringLiteralToken StringLiteral = (StringLiteralToken)Start;
                    return new StringLiteralExpression(StringLiteral.Literal, StringLiteral.Span);

                case TokenType.DateLiteral:
                    DateLiteralToken DateLiteral = (DateLiteralToken)Start;
                    return new DateLiteralExpression(DateLiteral.Literal, DateLiteral.Span);

                default:
                    Debug.Fail("Unexpected.");
                    break;
            }

            return null;
        }

        private Expression ParseCastExpression()
        {
            Token Start = Read();
            IntrinsicType OperatorType;
            Expression Operand;
            Location LeftParenthesisLocation;
            Location RightParenthesisLocation;

            switch (Start.Type)
            {
                case TokenType.CBool:
                    OperatorType = IntrinsicType.Boolean;
                    break;

                case TokenType.CChar:
                    OperatorType = IntrinsicType.Char;
                    break;

                case TokenType.CCur:
                    OperatorType = IntrinsicType.Currency;
                    break;

                case TokenType.CDate:
                    OperatorType = IntrinsicType.Date;
                    break;

                case TokenType.CDbl:
                    OperatorType = IntrinsicType.Double;
                    break;

                case TokenType.CByte:
                    OperatorType = IntrinsicType.Byte;
                    break;

                case TokenType.CShort:
                    OperatorType = IntrinsicType.Short;
                    break;

                case TokenType.CInt:
                    OperatorType = IntrinsicType.Integer;
                    break;

                case TokenType.CLng:
                    OperatorType = IntrinsicType.Long;
                    break;

                case TokenType.CSng:
                    OperatorType = IntrinsicType.Single;
                    break;

                case TokenType.CStr:
                    OperatorType = IntrinsicType.String;
                    break;

                case TokenType.CDec:
                    OperatorType = IntrinsicType.Decimal;
                    break;

                case TokenType.CVar:
                case TokenType.CObj:
                    OperatorType = IntrinsicType.Object;
                    break;

                case TokenType.CSByte:
                    OperatorType = IntrinsicType.SByte;
                    break;

                case TokenType.CUShort:
                    OperatorType = IntrinsicType.UShort;
                    break;

                case TokenType.CUInt:
                    OperatorType = IntrinsicType.UInteger;
                    break;

                case TokenType.CULng:
                    OperatorType = IntrinsicType.ULong;
                    break;

                default:
                    Debug.Fail("Unexpected.");
                    return null;
            }

            LeftParenthesisLocation = VerifyExpectedToken(TokenType.LeftParenthesis);
            Operand = ParseExpression();
            RightParenthesisLocation = VerifyExpectedToken(TokenType.RightParenthesis);

            return new IntrinsicCastExpression(OperatorType, LeftParenthesisLocation, Operand, RightParenthesisLocation, SpanFrom(Start));
        }

        private Expression ParseCastTypeExpression()
        {
            Token Start = Read();
            Expression Operand;
            TypeName Target;
            Location LeftParenthesisLocation;
            Location RightParenthesisLocation;
            Location CommaLocation;

            Debug.Assert(Start.Type == TokenType.CType || Start.Type == TokenType.DirectCast || Start.Type == TokenType.TryCast);

            LeftParenthesisLocation = VerifyExpectedToken(TokenType.LeftParenthesis);
            Operand = ParseExpression();

            if (ErrorInConstruct)
                ResyncAt(TokenType.Comma, TokenType.RightParenthesis);
            
            CommaLocation = VerifyExpectedToken(TokenType.Comma);
            Target = ParseTypeName(true);
            RightParenthesisLocation = VerifyExpectedToken(TokenType.RightParenthesis);

            if (Start.Type == TokenType.CType)
                return new CTypeExpression(LeftParenthesisLocation, Operand, CommaLocation, Target, RightParenthesisLocation, SpanFrom(Start));
            else if (Start.Type == TokenType.DirectCast)
                return new DirectCastExpression(LeftParenthesisLocation, Operand, CommaLocation, Target, RightParenthesisLocation, SpanFrom(Start));
            else
                return new TryCastExpression(LeftParenthesisLocation, Operand, CommaLocation, Target, RightParenthesisLocation, SpanFrom(Start));
        }

        private Expression ParseInstanceExpression()
        {
            Token Start = Read();
            InstanceType InstanceType = InstanceType.Me;

            switch (Start.Type)
            {
                case TokenType.Me:
                    InstanceType = InstanceType.Me;
                    break;

                case TokenType.MyClass:
                    InstanceType = InstanceType.MyClass;

                    if (Peek().Type != TokenType.Period)
                    {
                        ReportSyntaxError(SyntaxErrorType.ExpectedPeriodAfterMyClass, Start);
                    }

                    break;

                case TokenType.MyBase:
                    InstanceType = InstanceType.MyBase;

                    if (Peek().Type != TokenType.Period)
                    {
                        ReportSyntaxError(SyntaxErrorType.ExpectedPeriodAfterMyBase, Start);
                    }

                    break;

                default:
                    Debug.Fail("Unexpected.");
                    break;
            }

            return new InstanceExpression(InstanceType, Start.Span);
        }

        private Expression ParseParentheticalExpression()
        {
            Expression Operand;
            Token Start = Read();
            Location RightParenthesisLocation;

            Debug.Assert(Start.Type == TokenType.LeftParenthesis);

            Operand = ParseExpression();
            RightParenthesisLocation = VerifyExpectedToken(TokenType.RightParenthesis);

            return new ParentheticalExpression(Operand, RightParenthesisLocation, SpanFrom(Start));
        }

        private Expression ParseSimpleNameExpression()
        {
            Token Start = Peek();
            return new SimpleNameExpression(ParseSimpleName(false), SpanFrom(Start));
        }

        private Expression ParseDotBangExpression(Token start, Expression terminal)
        {
            SimpleName Name;
            Token DotBang;

            Debug.Assert(Peek().Type == TokenType.Period || Peek().Type == TokenType.Exclamation);

            DotBang = Read();
            Name = ParseSimpleName(true);

            if (DotBang.Type == TokenType.Period)
                return new QualifiedExpression(terminal, DotBang.Span.Start, Name, SpanFrom(start));
            else
                return new DictionaryLookupExpression(terminal, DotBang.Span.Start, Name, SpanFrom(start));
        }

        private Expression ParseCallOrIndexExpression(Token start, Expression terminal)
        {
            ArgumentCollection Arguments;

            // Because parentheses are used for array indexing, parameter passing, and array
            // declaring (via the Redim statement), there is some ambiguity about how to handle
            // a parenthesized list that begins with an expression. The most general case is to
            // parse it as an argument list.
            Arguments = ParseArguments();
            
            return new CallOrIndexExpression(terminal, Arguments, SpanFrom(start));
        }

        private Expression ParseTypeOfExpression()
        {
            Token Start = Peek();
            Expression Value;
            TypeName Target;
            Location IsLocation;

            Debug.Assert(Start.Type == TokenType.TypeOf);

            Read();
            Value = ParseBinaryOperatorExpression(PrecedenceLevel.Relational);

            if (ErrorInConstruct)
                ResyncAt(TokenType.Is);
            
            IsLocation = VerifyExpectedToken(TokenType.Is);
            Target = ParseTypeName(true);

            return new TypeOfExpression(Value, IsLocation, Target, SpanFrom(Start));
        }

        private Expression ParseGetTypeExpression()
        {
            Token Start = Read();
            TypeName Target;
            Location LeftParenthesisLocation;
            Location RightParenthesisLocation;

            Debug.Assert(Start.Type == TokenType.GetType);
            LeftParenthesisLocation = VerifyExpectedToken(TokenType.LeftParenthesis);
            Target = ParseTypeName(true, true);
            RightParenthesisLocation = VerifyExpectedToken(TokenType.RightParenthesis);

            return new GetTypeExpression(LeftParenthesisLocation, Target, RightParenthesisLocation, SpanFrom(Start));
        }

        private Expression ParseNewExpression()
        {
            Token Start = Read();
            Token TypeStart;
            TypeName Type;
            ArgumentCollection Arguments;
            Token ArgumentsStart;

            Debug.Assert(Start.Type == TokenType.New);

            TypeStart = Peek();
            Type = ParseTypeName(false);

            if (ErrorInConstruct)
                ResyncAt(TokenType.LeftParenthesis);
            
            ArgumentsStart = Peek();

            // This is an ambiguity in the grammar between
            //
            // New <Type> ( <Arguments> )
            // New <Type> <ArrayDeclaratorList> <AggregateInitializer>
            //
            // Try it as the first form, and if this fails, try the second.
            // (All valid instances of the second form have a beginning that is a valid
            // instance of the first form, so no spurious errors should result.)

            Arguments = ParseArguments();

            if ((Peek().Type == TokenType.LeftCurlyBrace || Peek().Type == TokenType.LeftParenthesis) && Arguments != null)
            {
                ArrayTypeName ArrayType;

                // Treat this as the form of New expression that allocates an array.
                GoBack(ArgumentsStart);
                ArrayType = ParseArrayTypeName(TypeStart, Type, true, false);

                if (Peek().Type == TokenType.LeftCurlyBrace)
                {
                    AggregateInitializer Initializer = ParseAggregateInitializer();
                    return new NewAggregateExpression(ArrayType, Initializer, SpanFrom(Start));
                }
                else
                {
                    HandleUnexpectedToken(TokenType.LeftCurlyBrace);
                }
            }

            return new NewExpression(Type, Arguments, SpanFrom(Start));
        }

        private Expression ParseTerminalExpression()
        {
            Token Start = Peek();
            Expression Terminal = null;

            Terminal = StartTerminalExpression(Start, Terminal);
            Terminal = FinishTerminalExpression(Start, Terminal);

            return Terminal;
        }

        private Expression StartTerminalExpression(Token Start, Expression Terminal)
        {
            switch (Start.Type)
            {
                case TokenType.Identifier:
                    Terminal = ParseSimpleNameExpression();
                    break;

                case TokenType.IntegerLiteral:
                case TokenType.FloatingPointLiteral:
                case TokenType.DecimalLiteral:
                case TokenType.CharacterLiteral:
                case TokenType.StringLiteral:
                case TokenType.DateLiteral:
                case TokenType.True:
                case TokenType.False:
                    Terminal = ParseLiteralExpression();
                    break;

                case TokenType.CBool:
                case TokenType.CByte:
                case TokenType.CShort:
                case TokenType.CInt:
                case TokenType.CLng:
                case TokenType.CDec:
                case TokenType.CSng:
                case TokenType.CDbl:
                case TokenType.CChar:
                case TokenType.CCur:
                case TokenType.CStr:
                case TokenType.CDate:
                case TokenType.CObj:
                case TokenType.CSByte:
                case TokenType.CUShort:
                case TokenType.CUInt:
                case TokenType.CULng:
                case TokenType.CVar:
                    Terminal = ParseCastExpression();
                    break;

                case TokenType.DirectCast:
                case TokenType.CType:
                case TokenType.TryCast:
                    Terminal = ParseCastTypeExpression();
                    break;

                case TokenType.Me:
                case TokenType.MyBase:
                case TokenType.MyClass:
                    Terminal = ParseInstanceExpression();
                    break;

                case TokenType.Global:
                    Terminal = new GlobalExpression(Read().Span);

                    if (Peek().Type != TokenType.Period)
                        ReportSyntaxError(SyntaxErrorType.ExpectedPeriodAfterGlobal, Start);

                    break;

                case TokenType.Nothing:
                    Terminal = new NothingExpression(Read().Span);
                    break;

                case TokenType.LeftParenthesis:
                    Terminal = ParseParentheticalExpression();
                    break;

                case TokenType.Period:
                case TokenType.Exclamation:
                    Terminal = ParseDotBangExpression(Start, null);
                    break;

                case TokenType.TypeOf:
                    Terminal = ParseTypeOfExpression();
                    break;

                case TokenType.GetType:
                    Terminal = ParseGetTypeExpression();
                    break;

                case TokenType.New:
                    Terminal = ParseNewExpression();
                    break;

                case TokenType.Short:
                case TokenType.Integer:
                case TokenType.Long:
                case TokenType.Decimal:
                case TokenType.Single:
                case TokenType.Double:
                case TokenType.Byte:
                case TokenType.Boolean:
                case TokenType.Char:
                case TokenType.Date:
                case TokenType.String:
                case TokenType.Object:
                case TokenType.Variant:
                    TypeName ReferencedType = ParseTypeName(false);

                    Terminal = new TypeReferenceExpression(ReferencedType, ReferencedType.Span);

                    if (Scanner.Peek().Type == TokenType.Period)
                        Terminal = ParseDotBangExpression(Start, Terminal);
                    else
                        HandleUnexpectedToken(TokenType.Period);

                    break;
                default:
                    ReportSyntaxError(SyntaxErrorType.ExpectedExpression, Peek());
                    Terminal = Expression.GetBadExpression(SpanFrom(Peek(), Peek()));
                    break;
            }

            return Terminal;
        }

        private Expression FinishTerminalExpression(Token Start, Expression Terminal)
        {
            Token Current = Peek();
            // Valid suffixes are ".", "!", and "(". Everything else is considered
            // to end the term.
            while (true)
            {
                bool IsDotBang = IsDotBangExpression();
                bool IsCallOrIndex = IsCallOrIndexExpression();
                bool IsGeneric = IsGenericExpression();

                if (IsDotBang)
                    Terminal = ParseDotBangExpression(Start, Terminal);
                else if (IsGeneric)
                    Terminal = ParseGenericQualifiedExpression(Start, Terminal);
                else if (IsCallOrIndex)
                    Terminal = ParseCallOrIndexExpression(Start, Terminal);
                else
                    break;
            }
            
            return Terminal;
        }

        private bool CanBypassParenthesis(Token Identifier)
        {
            bool IsCurrentIdentifier = IsIdentifier(Identifier);
            bool HasCall = HasCallOnStatement();
            bool Result = false;
            
            if (HasCall)
            {
                IsSubCall = true;
                HasParenthesisOnSubCall = true;
                HasCallToken = true;
            }
            else if (IsCurrentIdentifier && ImplementsVB60)
            {
                Token First = GetFirstTokenOfStatement(Identifier);
                Token NextToken = Next(Identifier);
                Token PreviousToken = Previous(Identifier);

                bool IsFirstIdentifier = IsIdentifier(First);
                bool IsFirstDotBang = IsDotBangToken(First);
                bool IsPotentialCaller = IsFirstIdentifier || IsFirstDotBang;

                bool HasSpaceAfter = ((IdentifierToken)Identifier).HasSpaceAfter;
                bool IsNextQualified = HasSpaceAfter && (NextToken != null && NextToken.Type == TokenType.Period);
                bool IsNextEvaluable = IsEvaluable(NextToken);
                bool IsNextNullArgument = (NextToken != null && NextToken.Type == TokenType.Comma);

                bool IsNextArgument = IsNextEvaluable || IsNextNullArgument || IsNextQualified;
                bool IsNextLeftParenthesis = NextToken != null && NextToken.Type == TokenType.LeftParenthesis;

                bool IsAssigmentStatement = CurrentBlockContextType() == TreeType.AssignmentStatement;
                bool IsParenthicalExpression = PreviousToken != null && PreviousToken.Type == TokenType.LeftParenthesis;

                if (IsPotentialCaller && !IsAssigmentStatement && !IsParenthicalExpression)
                {
                    if (!IsSubCall)
                    {
                        IsSubCall = IsNextLeftParenthesis || IsNextArgument;
                        HasParenthesisOnSubCall = IsNextLeftParenthesis;
                    }
                    Result = IsNextArgument;
                }
            }
            
            return Result;
        }

        private bool Legacy_CanBypassParenthesis(Token Identifier)
        {
            bool IsCurrentIdentifier = IsIdentifier(Identifier);
            bool HasCall = HasCallOnStatement();
            bool Result = false;

            if (HasCall)
            {
                IsSubCall = true;
                HasParenthesisOnSubCall = true;
                HasCallToken = true;
            }
            else if (IsCurrentIdentifier && ImplementsVB60)
            {
                Token First = GetFirstTokenOfStatement(Identifier);
                Token NextToken = Next(Identifier);
                Token PreviousToken = Previous(Identifier);

                bool IsFirstIdentifier = IsIdentifier(First);
                bool IsFirstDotBang = IsDotBangToken(First);
                bool IsPotentialCaller = IsFirstIdentifier || IsFirstDotBang;

                bool IsNextEvaluable = IsEvaluable(NextToken);
                bool IsNextNullArgument = (NextToken != null && NextToken.Type == TokenType.Comma);
                bool IsNextArgument = IsNextEvaluable || IsNextNullArgument;
                bool IsNextLeftParenthesis = NextToken != null && NextToken.Type == TokenType.LeftParenthesis;

                bool IsAssigmentStatement = CurrentBlockContextType() == TreeType.AssignmentStatement;
                bool IsParenthicalExpression = PreviousToken != null && PreviousToken.Type == TokenType.LeftParenthesis;

                if (!IsSubCall)
                {
                    if (!IsAssigmentStatement && !IsParenthicalExpression)
                    {
                        if (IsPotentialCaller)
                        {
                            if (IsNextArgument)
                            {
                                IsSubCall = true;
                                HasParenthesisOnSubCall = false;
                                Result = true;
                            }
                            else if (IsNextLeftParenthesis)
                            {
                                IsSubCall = true;
                                HasParenthesisOnSubCall = true;
                            }
                        }
                    }
                }
            }

            return Result;
        }

        private bool IsCallOrIndexExpression()
        {
            Token Current = Peek();
            Token Method = Previous();
            
            //bool UseParenthesis = UseParenthesisOnStatement(Method);
            bool BypassParenthesis = CanBypassParenthesis(Method);

            bool IsPreviousIdentifier = IsIdentifier(Method);
            bool IsArgument = IsBeginOfArgumentList(Current, BypassParenthesis);
            bool IsMethodArgument = IsPreviousIdentifier && IsArgument;

            bool IsLeftParenthesis = Current != null && Current.Type == TokenType.LeftParenthesis;
            bool IsPreviousRightParenthesis = Method != null && Method.Type == TokenType.RightParenthesis;
            bool IsArrayArgument = IsPreviousRightParenthesis && IsLeftParenthesis;
            
            return IsMethodArgument || IsArrayArgument;
        }

        private bool IsGenericExpression()
        {
            Token Current = Peek();
            Token First = PeekAheadOne();
            Move(Current);

            bool Result = (Current.Type == TokenType.LeftParenthesis && First != null && First.Type == TokenType.Of);

            return Result;
        }

        private bool IsDotBangExpression()
        {
            Token Current = Peek();
            Token PreviousToken = Previous(false);
            
            return !HasSpaceAfter(PreviousToken) && IsDotBangToken(Current);
        }

        private bool HasSpaceAfter(Token token)
        {
            if (token is IdentifierToken)
                return ((IdentifierToken)token).HasSpaceAfter;

            return false;
        }

        private GenericQualifiedExpression ParseGenericQualifiedExpression(Token Start, Expression Terminal)
        {
            Token LeftParenthesis = Read();
            TypeArgumentCollection ExpressionArguments = ParseTypeArguments(LeftParenthesis, false);
            Span SpanExpression = SpanFrom(Start);
            GenericQualifiedExpression GenericExpression = new GenericQualifiedExpression(Terminal, ExpressionArguments, SpanExpression);
            return GenericExpression;
        }

        private Expression ParseUnaryOperatorExpression()
        {
            Expression Operand = null;
            Expression expression = null;

            Token Start = Peek();
            UnaryOperatorType OperatorType = GetUnaryOperator(Start.Type);

            /*
            switch (Start.Type)
            {
                case TokenType.Minus:
                    OperatorType = UnaryOperatorType.Negate;
                    break;

                case TokenType.Plus:
                    OperatorType = UnaryOperatorType.UnaryPlus;
                    break;

                case TokenType.Not:
                    OperatorType = UnaryOperatorType.Not;
                    break;

                case TokenType.AddressOf:
                    Read();
                    Operand = ParseBinaryOperatorExpression(PrecedenceLevel.None);
                    return new AddressOfExpression(Operand, SpanFrom(Start, Operand));

                default:
                    return ParseTerminalExpression();
            }

            Read();
            Operand = ParseBinaryOperatorExpression(GetOperatorPrecedence(OperatorType));
            return new UnaryOperatorExpression(OperatorType, Operand, SpanFrom(Start, Operand));
            */

            if (OperatorType != UnaryOperatorType.None)
            {
                Read();
                Operand = ParseBinaryOperatorExpression(GetOperatorPrecedence(OperatorType));
                expression = new UnaryOperatorExpression(OperatorType, Operand, SpanFrom(Start, Operand));
            }
            else if (Start.Type == TokenType.AddressOf)
            {
                Read();
                Operand = ParseBinaryOperatorExpression(PrecedenceLevel.None);
                expression = new AddressOfExpression(Operand, SpanFrom(Start, Operand));
            }
            else
            {
                expression = ParseTerminalExpression();
            }

            return expression;
        }

        private Expression ParseBinaryOperatorExpression(PrecedenceLevel pendingPrecedence)
        {
            return ParseBinaryOperatorExpression(pendingPrecedence, false);
        }

        private Expression ParseBinaryOperatorExpression(PrecedenceLevel pendingPrecedence, bool allowRange)
        {
            Expression Expression;
            Token Start = Peek();

            Expression = ParseUnaryOperatorExpression();

            // Parse operators that follow the term according to precedence.
            while (true)
            {
                Expression Right = null;
                PrecedenceLevel Precedence;
                Location OperatorLocation;

                Token Current = Peek();
                Token NextToken = Next(Current);
                BinaryOperatorType OperatorType = GetBinaryOperator(Current.Type, allowRange);
                UnaryOperatorType NextOperator = GetUnaryOperator(NextToken.Type);

                if (ImplementsVB60 && OperatorType == BinaryOperatorType.Is && NextOperator == UnaryOperatorType.Not)
                {
                    OperatorType = BinaryOperatorType.IsNot;
                    Read();
                }

                if (OperatorType == BinaryOperatorType.None)
                    break;

                Precedence = GetOperatorPrecedence(OperatorType);

                // Only continue parsing if precedence is high enough
                if (Precedence <= pendingPrecedence)
                    break;

                OperatorLocation = ReadLocation();
                Right = ParseBinaryOperatorExpression(Precedence);
                Expression = new BinaryOperatorExpression(Expression, OperatorType, OperatorLocation, Right, SpanFrom(Start, Right));
            }
            
            return Expression;
        }

        private ExpressionCollection ParseExpressionList()
        {
            return ParseExpressionList(false);
        }

        private ExpressionCollection ParseExpressionList(bool requireExpression)
        {
            Token Start = Peek();
            List<Expression> Expressions = new List<Expression>();
            List<Location> CommaLocations = new List<Location>();

            if (CanEndStatement(Start) && !requireExpression)
            {
                return null;
            }

            do
            {
                if (Expressions.Count > 0)
                    CommaLocations.Add(ReadLocation());

                Expressions.Add(ParseExpression());

                if (ErrorInConstruct)
                    ResyncAt(TokenType.Comma);
            }
            while (Peek().Type == TokenType.Comma);

            if (Expressions.Count == 0 && CommaLocations.Count == 0)
            {
                return null;
            }
            else
            {
                return new ExpressionCollection(Expressions, CommaLocations, SpanFrom(Start));
            }
        }

        private Expression ParseExpression()
        {
            return ParseExpression(false);
        }

        private Expression ParseExpression(bool allowRange)
        {
            return ParseBinaryOperatorExpression(PrecedenceLevel.None, allowRange);
        }

        #endregion

        #region Statements

        private ExpressionStatement ParseExpressionStatement(TreeType type, bool operandOptional)
        {
            Token Start = Peek();
            Expression Operand = null;

            Read();

            if (!operandOptional || !CanEndStatement(Peek()))
                Operand = ParseExpression();

            if (ErrorInConstruct)
                ResyncAt();

            List<Comment> Comments = ParseTrailingComments();
            Span StartSpan = SpanFrom(Start);

            switch (type)
            {
                case TreeType.ReturnStatement:
                    if (ImplementsVB60 && Operand != null)
                    {
                        ReportSyntaxError(SyntaxErrorType.ExpectedEndOfStatement, Operand.Span);
                        Operand = null;
                    }
                    return new ReturnStatement(Operand, StartSpan, Comments, ImplementsVB60);

                case TreeType.ErrorStatement:
                    return new ErrorStatement(Operand, StartSpan, Comments);

                case TreeType.ThrowStatement:
                    return new ThrowStatement(Operand, StartSpan, Comments);

                default:
                    Debug.Fail("Unexpected!");
                    return null;
            }
        }

        // Parse a reference to a label, which can be an identifier or a line number.
        private void ParseLabelReference(ref SimpleName name, ref bool isLineNumber)
        {
            Token Start = Peek();

            if (Start.Type == TokenType.Identifier)
            {
                name = ParseSimpleName(false);
                isLineNumber = false;
            }
            else if (Start.Type == TokenType.IntegerLiteral)
            {
                IntegerLiteralToken IntegerLiteral = (IntegerLiteralToken)Start;

                if (IntegerLiteral.TypeCharacter != TypeCharacter.None)
                    ReportSyntaxError(SyntaxErrorType.SyntaxError, Start);

                name = new SimpleName(IntegerLiteral.Literal.ToString(), TypeCharacter.None, false, IntegerLiteral.Span);
                isLineNumber = true;

                bool HasColon = false;  
                if (ImplementsVB60 && IsBeginOfStatement(Start))
                {
                    Token NextToken = Next(Start);
                    if (NextToken.Type == TokenType.Colon)
                        HasColon = true;
                }
                else
                {
                    HasColon = true;
                }

                if (HasColon)
                    Read();
            }
            else
            {
                ReportSyntaxError(SyntaxErrorType.ExpectedIdentifier, Start);
                name = SimpleName.GetBadSimpleName(SpanFrom(Start));
                isLineNumber = false;
            }
        }

        private Statement ParseGotoStatement()
        {
            Token Start = Peek();
            SimpleName Name = null;
            bool IsLineNumber = false;

            Read();
            ParseLabelReference(ref Name, ref IsLineNumber);

            if (ErrorInConstruct)
                ResyncAt();

            return new GotoStatement(Name, IsLineNumber, SpanFrom(Start), ParseTrailingComments());
        }

        private Statement ParseGoSubStatement()
        {
            if (!ImplementsVB60)
            {
                return ParseGotoStatement();
            }
                        
            Token Start = Peek();
            SimpleName Name = null;
            bool IsLineNumber = false;

            Read();
            ParseLabelReference(ref Name, ref IsLineNumber);

            if (ErrorInConstruct)
                ResyncAt();

            return new GoSubStatement(Name, IsLineNumber, SpanFrom(Start), ParseTrailingComments());
        }

        private Statement ParseContinueStatement()
        {
            Token Start = Peek();
            BlockType ContinueType;
            Location ContinueArgumentLocation = Location.Empty;

            Read();

            ContinueType = GetContinueType(Peek().Type);

            if (ContinueType == BlockType.None)
            {
                ReportSyntaxError(SyntaxErrorType.ExpectedContinueKind, Peek());
                ResyncAt();
            }
            else
            {
                ContinueArgumentLocation = ReadLocation();
            }

            return new ContinueStatement(ContinueType, ContinueArgumentLocation, SpanFrom(Start), ParseTrailingComments());
        }

        private Statement ParseExitStatement()
        {
            Token Start = Peek();
            BlockType ExitType;
            Location ExitArgumentLocation = Location.Empty;

            Read();

            ExitType = GetExitType(Peek().Type);

            if (ExitType == BlockType.None)
            {
                ReportSyntaxError(SyntaxErrorType.ExpectedExitKind, Peek());
                ResyncAt();
            }
            else
            {
                ExitArgumentLocation = ReadLocation();
            }

            return new ExitStatement(ExitType, ExitArgumentLocation, SpanFrom(Start), ParseTrailingComments());
        }

        private Statement ParseEndStatement()
        {
            Token Start = Peek();

            if (Start.Type != TokenType.Wend)
                Start = Read();
            
            BlockType EndType = GetBlockType(Peek().Type);

            if (EndType == BlockType.None)
                return new EndStatement(SpanFrom(Start), ParseTrailingComments());
            
            return new EndBlockStatement(EndType, ReadLocation(), SpanFrom(Start), ParseTrailingComments());
        }

        private Statement ParseRaiseEventStatement()
        {
            Token Start = Peek();
            SimpleName Name;
            ArgumentCollection Arguments;

            Read();
            Name = ParseSimpleName(true);

            if (ErrorInConstruct)
                ResyncAt();

            Arguments = ParseArguments();

            return new RaiseEventStatement(Name, Arguments, SpanFrom(Start), ParseTrailingComments());
        }

        private Statement ParseHandlerStatement()
        {
            Token Start = Peek();
            Expression EventName;
            Expression DelegateExpression;
            Location CommaLocation;

            Read();
            EventName = ParseExpression();

            if (ErrorInConstruct)
                ResyncAt(TokenType.Comma);

            CommaLocation = VerifyExpectedToken(TokenType.Comma);
            DelegateExpression = ParseExpression();

            if (ErrorInConstruct)
                ResyncAt();

            if (Start.Type == TokenType.AddHandler)
                return new AddHandlerStatement(EventName, CommaLocation, DelegateExpression, SpanFrom(Start), ParseTrailingComments());
            else
                return new RemoveHandlerStatement(EventName, CommaLocation, DelegateExpression, SpanFrom(Start), ParseTrailingComments());
        }

        private Statement ParseOnErrorStatement()
        {
            Token Start = Read();
            OnErrorType OnErrorType;
            Token NextToken;
            SimpleName Name = null;
            bool IsLineNumber = false;
            Location ErrorLocation = Location.Empty;
            Location ResumeOrGoToLocation = Location.Empty;
            Location NextOrZeroOrMinusLocation = Location.Empty;
            Location OneLocation = Location.Empty;

            if (Peek().Type == TokenType.Error || Peek().AsUnreservedKeyword() == TokenType.Error)
            {
                ErrorLocation = ReadLocation();
                NextToken = Peek();

                if (NextToken.Type == TokenType.Resume)
                {
                    ResumeOrGoToLocation = ReadLocation();

                    if (Peek().Type == TokenType.Next)
                        NextOrZeroOrMinusLocation = ReadLocation();
                    else
                        ReportSyntaxError(SyntaxErrorType.ExpectedNext, Peek());

                    OnErrorType = OnErrorType.Next;
                }
                else if (NextToken.Type == TokenType.GoTo)
                {
                    ResumeOrGoToLocation = ReadLocation();
                    NextToken = Peek();

                    if (NextToken.Type == TokenType.IntegerLiteral && ((IntegerLiteralToken)NextToken).Literal == 0)
                    {
                        NextOrZeroOrMinusLocation = ReadLocation();
                        OnErrorType = OnErrorType.Zero;
                    }
                    else if (NextToken.Type == TokenType.Minus)
                    {
                        Token NextNextToken;

                        NextOrZeroOrMinusLocation = ReadLocation();
                        NextNextToken = Peek();

                        if (NextNextToken.Type == TokenType.IntegerLiteral && ((IntegerLiteralToken)NextNextToken).Literal == 1)
                        {
                            OneLocation = ReadLocation();
                            OnErrorType = OnErrorType.MinusOne;
                        }
                        else
                        {
                            GoBack(NextToken);

                            OnErrorType = OnErrorType.Label;
                            ParseLabelReference(ref Name, ref IsLineNumber);

                            if (ErrorInConstruct)
                                ResyncAt();
                        }
                    }
                    else
                    {
                        OnErrorType = OnErrorType.Label;
                        ParseLabelReference(ref Name, ref IsLineNumber);

                        if (ErrorInConstruct)
                            ResyncAt();
                    }
                }
                else
                {
                    ReportSyntaxError(SyntaxErrorType.ExpectedResumeOrGoto, Peek());
                    OnErrorType = OnErrorType.Bad;
                    ResyncAt();
                }
            }
            else
            {
                ReportSyntaxError(SyntaxErrorType.ExpectedError, Peek());
                OnErrorType = OnErrorType.Bad;
                ResyncAt();
            }

            return new OnErrorStatement(OnErrorType, ErrorLocation, ResumeOrGoToLocation, NextOrZeroOrMinusLocation, OneLocation, Name, IsLineNumber, SpanFrom(Start), ParseTrailingComments());
        }

        private Statement ParseResumeStatement()
        {
            Token Start = Read();
            ResumeType ResumeType = ResumeType.None;
            SimpleName Name = null;
            bool IsLineNumber = false;
            Location NextLocation = Location.Empty;

            if (!CanEndStatement(Peek()))
            {
                if (Peek().Type == TokenType.Next)
                {
                    NextLocation = ReadLocation();
                    ResumeType = ResumeType.Next;
                }
                else
                {
                    ParseLabelReference(ref Name, ref IsLineNumber);

                    if (ErrorInConstruct)
                        ResumeType = ResumeType.None;
                    else
                        ResumeType = ResumeType.Label;
                }
            }

            return new ResumeStatement(ResumeType, NextLocation, Name, IsLineNumber, SpanFrom(Start), ParseTrailingComments());
        }

        private Statement ParseReDimStatement()
        {
            Token Start = Read();
            Location PreserveLocation = Location.Empty;
            ExpressionCollection Variables = null;

            if (Peek().AsUnreservedKeyword() == TokenType.Preserve)
                PreserveLocation = ReadLocation();

            Variables = ParseExpressionList(true);
            
            return new ReDimStatement(PreserveLocation, Variables, SpanFrom(Start), ParseTrailingComments());
        }

        private Statement ParseEraseStatement()
        {
            Token Start = Read();
            ExpressionCollection Variables = ParseExpressionList(true);

            return new EraseStatement(Variables, SpanFrom(Start), ParseTrailingComments());
        }

        private Statement ParseAssignmentStatement()
        {
            Token Start = Peek();
            Token Acessor = null;
            Statement Result = null;

            bool IsOperator = IdentifierToken.IsAssignmentAccessor(Start.Type);
            if (ImplementsVB60 && IsOperator)
                Acessor = Read();

            Expression Target = ParseBinaryOperatorExpression(PrecedenceLevel.Power);
            if (ErrorInConstruct)
                ResyncAt(TokenType.Equals);

            if (GetAssignmentOperator(Peek().Type) != TreeType.SyntaxError)
                Result = ParseAssignmentStatement(Target, Acessor);

            return Result;
        }

        private Statement ParseAssignmentStatement(Expression target)
        {
            return ParseAssignmentStatement(target, null);
        }

        private Statement ParseAssignmentStatement(Expression target, Token acessor)
        {
            PushBlockContext(TreeType.AssignmentStatement);

            BinaryOperatorType CompoundOperator;
            Expression Source;
            Token Operator;
            Statement Result;

            Operator = Read();
            CompoundOperator = GetCompoundAssignmentOperatorType(Operator.Type);
            Source = ParseExpression();

            if (ErrorInConstruct)
                ResyncAt();

            IList<Comment> comments = ParseTrailingComments();
            Location start = Operator.Span.Start;
            Span span = SpanFrom(target.Span.Start);

            if (CompoundOperator == (BinaryOperatorType)TreeType.SyntaxError)
                Result = new AssignmentStatement(target, start, Source, span, comments, acessor);
            else
                Result = new CompoundAssignmentStatement(CompoundOperator, target, start, Source, span, comments);

            PopBlockContext();

            return Result;
        }

        private Statement ParseCallStatement()
        {
            return ParseCallStatement(null);
        }

        private Statement ParseCallStatement(Expression target)
        {
            Token Start = Peek();
            Location StartLocation = Location.Empty;
            Location CallLocation = Location.Empty;
            ArgumentCollection Arguments = null;

            if (target == null)
            {
                StartLocation = Start.Span.Start;

                if (Start.Type == TokenType.Call)
                    CallLocation = ReadLocation();

                target = ParseExpression();

                if (ErrorInConstruct)
                    ResyncAt();
            }
            else
            {
                StartLocation = target.Span.Start;
            }

            if (target.Type == TreeType.CallOrIndexExpression)
            {
                // Extract the operands of the call/index expression and make
                // them operands of the call statement.
                CallOrIndexExpression CallOrIndexExpression = (CallOrIndexExpression)target;

                target = CallOrIndexExpression.TargetExpression;
                Arguments = CallOrIndexExpression.Arguments;
            }

            return new CallStatement(CallLocation, target, Arguments, SpanFrom(StartLocation), ParseTrailingComments());
        }

        private Statement ParseMidAssignmentStatement()
        {
            Token Start = Read();
            IdentifierToken Identifier = (IdentifierToken)Start;
            bool HasTypeCharacter = false;
            Location LeftParenthesisLocation = Location.Empty;
            Expression Target = null;
            Location StartCommaLocation = Location.Empty;
            Expression StartExpression = null;
            Location LengthCommaLocation = Location.Empty;
            Expression LengthExpression = null;
            Location RightParenthesisLocation = Location.Empty;
            Location OperatorLocation = Location.Empty;
            Expression Source = null;

            if (Identifier.TypeCharacter == TypeCharacter.StringSymbol)
                HasTypeCharacter = true;
            else if (Identifier.TypeCharacter != TypeCharacter.None)
                goto NotMidAssignment;

            if (Peek().Type == TokenType.LeftParenthesis)
                LeftParenthesisLocation = VerifyExpectedToken(TokenType.LeftParenthesis);
            else
                goto NotMidAssignment;

            // This is very unfortunate: ideally, we would continue parsing to
            // make sure the entire statement matched the form of a Mid assignment
            // statement. That way something like "Mid(10) = 5", where Mid is an
            // array identifier wouldn't cause an error. Alas, it's not that simple
            // because what about something that's in error? We could fall back on
            // error, but we have no way of backtracking on errors at the moment.
            // So we're going to do what the official compiler does: if we see
            // Mid and (, you've got yourself a Mid assignment statement!
            Target = ParseExpression();

            if (ErrorInConstruct)
                ResyncAt(TokenType.Comma);

            StartCommaLocation = VerifyExpectedToken(TokenType.Comma);
            StartExpression = ParseExpression();

            if (ErrorInConstruct)
                ResyncAt(TokenType.Comma, TokenType.RightParenthesis);

            if (Peek().Type == TokenType.Comma)
            {
                LengthCommaLocation = VerifyExpectedToken(TokenType.Comma);
                LengthExpression = ParseExpression();

                if (ErrorInConstruct)
                    ResyncAt(TokenType.RightParenthesis);
            }

            RightParenthesisLocation = VerifyExpectedToken(TokenType.RightParenthesis);
            OperatorLocation = VerifyExpectedToken(TokenType.Equals);

            Source = ParseExpression();

            if (ErrorInConstruct)
                ResyncAt();

            return new MidAssignmentStatement(HasTypeCharacter, LeftParenthesisLocation, Target, StartCommaLocation, StartExpression, LengthCommaLocation, LengthExpression, RightParenthesisLocation, OperatorLocation, Source,
            SpanFrom(Start), ParseTrailingComments());
        NotMidAssignment:

            GoBack(Start);
            return null;
        }

        private Statement ParseLocalDeclarationStatement()
        {
            Token Start = Peek();
            ModifierCollection Modifiers = null;
            const ModifierTypes ValidModifiers = ModifierTypes.Dim | ModifierTypes.Const | ModifierTypes.Static;

            Modifiers = ParseDeclarationModifierList();
            ValidateModifierList(Modifiers, ValidModifiers);

            if (Modifiers == null)
                ReportSyntaxError(SyntaxErrorType.ExpectedModifier, Peek());
            else if (Modifiers.Count > 1)
                ReportSyntaxError(SyntaxErrorType.InvalidVariableModifiers, Modifiers.Span);

            return new LocalDeclarationStatement(Modifiers, ParseVariableDeclarators(Modifiers), SpanFrom(Start), ParseTrailingComments());
        }

        private Statement ParseLabelStatement()
        {
            SimpleName Name = null;
            bool IsLineNumber = false;
            Token Start = Peek();

            ParseLabelReference(ref Name, ref IsLineNumber);
            return new LabelStatement(Name, IsLineNumber, SpanFrom(Start), ParseTrailingComments());
        }

        private Statement ParseExpressionBlockStatement(TreeType blockType)
        {
            Token Start = Read();
            Expression Expression = ParseExpression();
            StatementCollection StatementCollection;
            Statement EndStatement = null;
            List<Comment> Comments = null;

            if (ErrorInConstruct)
            {
                ResyncAt();
            }

            StatementCollection = ParseStatementBlock(SpanFrom(Start), blockType, ref Comments, ref EndStatement);

            switch (blockType)
            {
                case TreeType.WithBlockStatement:
                    return new WithBlockStatement(Expression, StatementCollection, (EndBlockStatement)EndStatement, SpanFrom(Start), Comments);

                case TreeType.SyncLockBlockStatement:
                    return new SyncLockBlockStatement(Expression, StatementCollection, (EndBlockStatement)EndStatement, SpanFrom(Start), Comments);

                case TreeType.WhileBlockStatement:
                    return new WhileBlockStatement(Expression, StatementCollection, (EndBlockStatement)EndStatement, SpanFrom(Start), Comments);

                default:
                    Debug.Fail("Unexpected!");
                    return null;
            }
        }

        private Statement ParseUsingBlockStatement()
        {
            Token Start = Read();
            Expression Expression = null;
            VariableDeclaratorCollection VariableDeclarators = null;
            StatementCollection StatementCollection;
            Statement EndStatement = null;
            List<Comment> Comments = null;
            Token NextToken = PeekAheadOne();

            if (NextToken.Type == TokenType.As || NextToken.Type == TokenType.Equals)
                VariableDeclarators = ParseVariableDeclarators(null);
            else
                Expression = ParseExpression();
            
            if (ErrorInConstruct)
                ResyncAt();
        
            StatementCollection = ParseStatementBlock(SpanFrom(Start), TreeType.UsingBlockStatement, ref Comments, ref EndStatement);

            if (Expression == null)
                return new UsingBlockStatement(VariableDeclarators, StatementCollection, (EndBlockStatement)EndStatement, SpanFrom(Start), Comments);
        
            else
                return new UsingBlockStatement(Expression, StatementCollection, (EndBlockStatement)EndStatement, SpanFrom(Start), Comments);
        }

        private Expression ParseOptionalWhileUntilClause(ref bool isWhile, ref Location whileOrUntilLocation)
        {
            Expression Expression = null;

            if (!CanEndStatement(Peek()))
            {
                Token Token = Peek();

                if (Token.Type == TokenType.While || Token.AsUnreservedKeyword() == TokenType.Until)
                {
                    isWhile = (Token.Type == TokenType.While);
                    whileOrUntilLocation = ReadLocation();
                    Expression = ParseExpression();

                    if (ErrorInConstruct)
                        ResyncAt();
                }
                else
                {
                    ReportSyntaxError(SyntaxErrorType.SyntaxError, Peek());
                    ResyncAt();
                }
            }

            return Expression;
        }

        private Statement ParseDoBlockStatement()
        {
            Token Start = Read();
            bool IsWhile = false;
            Expression Expression = null;
            Location WhileOrUntilLocation = Location.Empty;
            StatementCollection StatementCollection = null;
            Statement EndStatement = null;
            LoopStatement LoopStatement = null;
            List<Comment> Comments = null;

            Expression = ParseOptionalWhileUntilClause(ref IsWhile, ref WhileOrUntilLocation);

            StatementCollection = ParseStatementBlock(SpanFrom(Start), TreeType.DoBlockStatement, ref Comments, ref EndStatement);
            LoopStatement = (LoopStatement)EndStatement;

            if (Expression != null && LoopStatement != null && LoopStatement.Expression != null)
                ReportSyntaxError(SyntaxErrorType.LoopDoubleCondition, LoopStatement.Expression.Span);

            return new DoBlockStatement(Expression, IsWhile, WhileOrUntilLocation, StatementCollection, LoopStatement, SpanFrom(Start), Comments);
        }

        private Statement ParseLoopStatement()
        {
            Token Start = Read();
            bool IsWhile = false;
            Expression Expression = null;
            Location WhileOrUntilLocation = Location.Empty;

            Expression = ParseOptionalWhileUntilClause(ref IsWhile, ref WhileOrUntilLocation);
            return new LoopStatement(Expression, IsWhile, WhileOrUntilLocation, SpanFrom(Start), ParseTrailingComments());
        }

        private Expression ParseForLoopControlVariable(ref VariableDeclarator variableDeclarator)
        {
            Token Start = Peek();

            if (Start.Type == TokenType.Identifier)
            {
                Token NextToken = PeekAheadOne();
                Expression expression = null;

                // CONSIDER: Should we just always parse this as a variable declaration?
                if (NextToken.Type == TokenType.As)
                {
                    if (ImplementsVB60)
                        ReportSyntaxError(SyntaxErrorType.ExpectedEquals, NextToken);

                    variableDeclarator = ParseForLoopVariableDeclarator(ref expression);
                    return expression;
                }
                else if (NextToken.Type == TokenType.LeftParenthesis)
                {
                    // CONSIDER: Only do this if the token previous to the As is a right parenthesis
                    if (PeekAheadFor(TokenType.As, TokenType.In, TokenType.Equals) == TokenType.As)
                    {
                        if (ImplementsVB60)
                            ReportSyntaxError(SyntaxErrorType.ExpectedEquals, NextToken);

                        variableDeclarator = ParseForLoopVariableDeclarator(ref expression);
                        return expression;
                    }
                }
            }

            return ParseBinaryOperatorExpression(PrecedenceLevel.Relational);
        }

        private Statement ParseForBlockStatement()
        {
            Token Start = Read();

            if (Peek().Type != TokenType.Each)
            {
                Expression ControlExpression = null;
                Expression LowerBoundExpression = null;
                Expression UpperBoundExpression = null;
                Expression StepExpression = null;
                Location EqualLocation = Location.Empty;
                Location ToLocation = Location.Empty;
                Location StepLocation = Location.Empty;
                VariableDeclarator VariableDeclarator = null;
                StatementCollection Statements = null;
                Statement NextStatement = null;
                List<Comment> Comments = null;

                ControlExpression = ParseForLoopControlVariable(ref VariableDeclarator);

                if (ErrorInConstruct)
                    ResyncAt(TokenType.Equals, TokenType.To);
                
                if (Peek().Type == TokenType.Equals)
                {
                    EqualLocation = ReadLocation();
                    LowerBoundExpression = ParseExpression();

                    if (ErrorInConstruct)
                    {
                        ResyncAt(TokenType.To);
                    }
                }
                else
                {
                    ReportSyntaxError(SyntaxErrorType.SyntaxError, Peek());
                    ResyncAt(TokenType.To);
                }

                if (Peek().Type == TokenType.To)
                {
                    ToLocation = ReadLocation();
                    UpperBoundExpression = ParseExpression();

                    if (ErrorInConstruct)
                    {
                        ResyncAt(TokenType.Step);
                    }
                }
                else
                {
                    ReportSyntaxError(SyntaxErrorType.SyntaxError, Peek());
                    ResyncAt(TokenType.Step);
                }

                if (Peek().Type == TokenType.Step)
                {
                    StepLocation = ReadLocation();
                    StepExpression = ParseExpression();

                    if (ErrorInConstruct)
                    {
                        ResyncAt();
                    }
                }

                Statements = ParseStatementBlock(SpanFrom(Start), TreeType.ForBlockStatement, ref Comments, ref NextStatement);

                return new ForBlockStatement(ControlExpression, VariableDeclarator, EqualLocation, LowerBoundExpression, ToLocation, UpperBoundExpression, StepLocation, StepExpression, Statements, (NextStatement)NextStatement,
                SpanFrom(Start), Comments);
            }
            else
            {
                Location EachLocation = Location.Empty;
                Expression ControlExpression = null;
                Location InLocation = Location.Empty;
                VariableDeclarator VariableDeclarator = null;
                Expression CollectionExpression = null;
                StatementCollection Statements = null;
                Statement NextStatement = null;
                List<Comment> Comments = null;

                EachLocation = ReadLocation();
                ControlExpression = ParseForLoopControlVariable(ref VariableDeclarator);

                if (ErrorInConstruct)
                {
                    ResyncAt(TokenType.In);
                }

                if (Peek().Type == TokenType.In)
                {
                    InLocation = ReadLocation();
                    CollectionExpression = ParseExpression();

                    if (ErrorInConstruct)
                    {
                        ResyncAt();
                    }
                }
                else
                {
                    ReportSyntaxError(SyntaxErrorType.SyntaxError, Peek());
                    ResyncAt();
                }

                Statements = ParseStatementBlock(SpanFrom(Start), TreeType.ForBlockStatement, ref Comments, ref NextStatement);

                return new ForEachBlockStatement(EachLocation, ControlExpression, VariableDeclarator, InLocation, CollectionExpression, Statements, (NextStatement)NextStatement, SpanFrom(Start), Comments);
            }
        }

        private Statement ParseNextStatement()
        {
            Token Start = Read();

            return new NextStatement(ParseExpressionList(), SpanFrom(Start), ParseTrailingComments());
        }

        private Statement ParseTryBlockStatement()
        {
            Token Start = Read();
            StatementCollection TryStatementList;
            StatementCollection StatementCollection;
            Statement EndBlockStatement = null;
            List<Statement> CatchBlocks = new List<Statement>();
            StatementCollection CatchBlockList = null;
            FinallyBlockStatement FinallyBlock = null;
            List<Comment> Comments = null;
            List<Comment> Empty = null;

            TryStatementList = ParseStatementBlock(SpanFrom(Start), TreeType.TryBlockStatement, ref Comments, ref EndBlockStatement);

            while ((EndBlockStatement != null) && (EndBlockStatement.Type != TreeType.EndBlockStatement))
            {
                if (EndBlockStatement.Type == TreeType.CatchStatement)
                {
                    CatchStatement CatchStatement = (CatchStatement)EndBlockStatement;

                    StatementCollection = ParseStatementBlock(CatchStatement.Span, TreeType.CatchBlockStatement, ref Empty, ref EndBlockStatement);
                    CatchBlocks.Add(new CatchBlockStatement(CatchStatement, StatementCollection, SpanFrom(CatchStatement, EndBlockStatement), null));
                }
                else
                {
                    FinallyStatement FinallyStatement = (FinallyStatement)EndBlockStatement;

                    StatementCollection = ParseStatementBlock(FinallyStatement.Span, TreeType.FinallyBlockStatement, ref Empty, ref EndBlockStatement);
                    FinallyBlock = new FinallyBlockStatement(FinallyStatement, StatementCollection, SpanFrom(FinallyStatement, EndBlockStatement), null);
                }
            }

            if (CatchBlocks.Count > 0)
            {
                CatchBlockList = new StatementCollection(CatchBlocks, null, new Span(((CatchBlockStatement)CatchBlocks[0]).Span.Start, ((CatchBlockStatement)CatchBlocks[CatchBlocks.Count - 1]).Span.Finish));
            }

            return new TryBlockStatement(TryStatementList, CatchBlockList, FinallyBlock, (EndBlockStatement)EndBlockStatement, SpanFrom(Start), Comments);
        }

        private Statement ParseCatchStatement()
        {
            Token Start = Read();
            SimpleName Name = null;
            Location AsLocation = Location.Empty;
            TypeName Type = null;
            Location WhenLocation = Location.Empty;
            Expression Filter = null;

            if (Peek().Type == TokenType.Identifier)
            {
                Name = ParseSimpleName(false);

                if (Peek().Type == TokenType.As)
                {
                    AsLocation = ReadLocation();
                    Type = ParseTypeName(false);

                    if (ErrorInConstruct)
                        ResyncAt(TokenType.When);
                }
            }

            if (Peek().Type == TokenType.When)
            {
                WhenLocation = ReadLocation();
                Filter = ParseExpression();
            }

            return new CatchStatement(Name, AsLocation, Type, WhenLocation, Filter, SpanFrom(Start), ParseTrailingComments());
        }

        private Statement ParseCaseStatement()
        {
            Token Start = Read();
            List<Location> CommaLocations;
            List<CaseClause> Cases;
            Token CasesStart = Peek();

            if (Peek().Type == TokenType.Else)
            {
                return new CaseElseStatement(ReadLocation(), SpanFrom(Start), ParseTrailingComments());
            }
            else
            {
                CommaLocations = new List<Location>();
                Cases = new List<CaseClause>();

                do
                {
                    if (Cases.Count > 0)
                    {
                        CommaLocations.Add(ReadLocation());
                    }

                    Cases.Add(ParseCase());
                }
                while (Peek().Type == TokenType.Comma);

                return new CaseStatement(new CaseClauseCollection(Cases, CommaLocations, SpanFrom(CasesStart)), SpanFrom(Start), ParseTrailingComments());
            }
        }

        private Statement ParseSelectBlockStatement()
        {
            Token Start = Read();
            Location CaseLocation = Location.Empty;
            Expression SelectExpression = null;
            StatementCollection Statements = null;
            Statement EndBlockStatement = null;
            List<Statement> CaseBlocks = new List<Statement>();
            StatementCollection CaseBlockList = null;
            CaseElseBlockStatement CaseElseBlockStatement = null;
            List<Comment> Comments = null;
            List<Comment> Empty = null;

            if (Peek().Type == TokenType.Case)
            {
                CaseLocation = ReadLocation();
            }

            SelectExpression = ParseExpression();

            if (ErrorInConstruct)
            {
                ResyncAt();
            }

            Statements = ParseStatementBlock(SpanFrom(Start), TreeType.SelectBlockStatement, ref Comments, ref EndBlockStatement);

            if (Statements != null && Statements.Count != 0)
            {
                foreach (Statement Statement in Statements)
                {
                    if (Statement.Type != TreeType.EmptyStatement)
                        ReportSyntaxError(SyntaxErrorType.ExpectedCase, Statements.Span);
                }
            }

            while ((EndBlockStatement != null) && (EndBlockStatement.Type != TreeType.EndBlockStatement))
            {
                Statement CaseStatement = EndBlockStatement;
                StatementCollection CaseStatements;

                if (CaseStatement.Type == TreeType.CaseStatement)
                {
                    CaseStatements = ParseStatementBlock(CaseStatement.Span, TreeType.CaseBlockStatement, ref Empty, ref EndBlockStatement);
                    CaseBlocks.Add(new CaseBlockStatement((CaseStatement)CaseStatement, CaseStatements, SpanFrom(CaseStatement, EndBlockStatement), null));
                }
                else
                {
                    CaseStatements = ParseStatementBlock(CaseStatement.Span, TreeType.CaseElseBlockStatement, ref Empty, ref EndBlockStatement);
                    CaseElseBlockStatement = new CaseElseBlockStatement((CaseElseStatement)CaseStatement, CaseStatements, SpanFrom(CaseStatement, EndBlockStatement), null);
                }
            }

            if (CaseBlocks.Count > 0)
            {
                CaseBlockList = new StatementCollection(CaseBlocks, null, new Span(((CaseBlockStatement)CaseBlocks[0]).Span.Start, ((CaseBlockStatement)CaseBlocks[CaseBlocks.Count - 1]).Span.Finish));
            }

            return new SelectBlockStatement(CaseLocation, SelectExpression, Statements, CaseBlockList, CaseElseBlockStatement, (EndBlockStatement)EndBlockStatement, SpanFrom(Start), Comments);
        }

        private Statement ParseElseIfStatement()
        {
            Token Start = Read();
            Location ThenLocation = Location.Empty;
            Expression Expression = null;

            Expression = ParseExpression();

            if (ErrorInConstruct)
            {
                ResyncAt(TokenType.Then);
            }

            if (Peek().Type == TokenType.Then)
            {
                ThenLocation = ReadLocation();
            }

            return new ElseIfStatement(Expression, ThenLocation, SpanFrom(Start), ParseTrailingComments());
        }

        private Statement ParseIfBlockStatement()
        {
            Token Start = Read();
            Expression Expression = null;
            Location ThenLocation = Location.Empty;
            StatementCollection Statements = null;
            StatementCollection IfStatements = null;
            Statement EndBlockStatement = null;
            List<Statement> ElseIfBlocks = new List<Statement>();
            StatementCollection ElseIfBlockList = null;
            ElseBlockStatement ElseBlockStatement = null;
            List<Comment> Comments = null;
            List<Comment> Empty = null;

            Expression = ParseExpression();

            if (ErrorInConstruct)
            {
                ResyncAt(TokenType.Then);
            }

            if (Peek().Type == TokenType.Then)
            {
                ThenLocation = ReadLocation();

                if (!CanEndStatement(Peek()))
                {
                    Location ElseLocation = Location.Empty;
                    StatementCollection ElseStatements = null;

                    // We're in a line If context
                    AtBeginningOfLine = false;
                    IfStatements = ParseLineIfStatementBlock();

                    if (Peek().Type == TokenType.Else)
                    {
                        ElseLocation = ReadLocation();
                        ElseStatements = ParseLineIfStatementBlock();
                    }

                    return new LineIfStatement(Expression, ThenLocation, IfStatements, ElseLocation, ElseStatements, SpanFrom(Start), ParseTrailingComments());
                }
            }

            IfStatements = ParseStatementBlock(SpanFrom(Start), TreeType.IfBlockStatement, ref Comments, ref EndBlockStatement);

            while ((EndBlockStatement != null) && (EndBlockStatement.Type != TreeType.EndBlockStatement))
            {
                Statement ElseStatement = EndBlockStatement;

                if (ElseStatement.Type == TreeType.ElseIfStatement)
                {
                    Statements = ParseStatementBlock(ElseStatement.Span, TreeType.ElseIfBlockStatement, ref Empty, ref EndBlockStatement);
                    ElseIfBlocks.Add(new ElseIfBlockStatement((ElseIfStatement)ElseStatement, Statements, SpanFrom(ElseStatement, EndBlockStatement), null));
                }
                else
                {
                    Statements = ParseStatementBlock(ElseStatement.Span, TreeType.ElseBlockStatement, ref Empty, ref EndBlockStatement);
                    ElseBlockStatement = new ElseBlockStatement((ElseStatement)ElseStatement, Statements, SpanFrom(ElseStatement, EndBlockStatement), null);
                }
            }

            if (ElseIfBlocks.Count > 0)
            {
                ElseIfBlockList = new StatementCollection(ElseIfBlocks, null, new Span(((Statement)ElseIfBlocks[0]).Span.Start, ((Statement)ElseIfBlocks[ElseIfBlocks.Count - 1]).Span.Finish));
            }

            return new IfBlockStatement(Expression, ThenLocation, IfStatements, ElseIfBlockList, ElseBlockStatement, (EndBlockStatement)EndBlockStatement, SpanFrom(Start), Comments);
        }

        private Statement ParseStatement()
        {
            Token empty = null;
            return ParseStatement(ref empty);
        }

        private Statement ParseStatement(ref Token terminator)
        {
            Token Start = Peek();
            Statement Statement = null;
            bool IsSyntaxError = false;
            
            if (AtBeginningOfLine)
            {
                while (ParsePreprocessorStatement(true))
                    Start = Peek();
            }

            ErrorInConstruct = false;

            switch (Start.Type)
            {
                case TokenType.GoTo:
                    Statement = ParseGotoStatement();
                    break;

                case TokenType.GoSub:
                    Statement = ParseGoSubStatement();
                    break;

                case TokenType.Exit:
                    Statement = ParseExitStatement();
                    break;

                case TokenType.Continue:
                    Statement = ParseContinueStatement();
                    break;

                case TokenType.Stop:
                    Statement = new StopStatement(SpanFrom(Read()), ParseTrailingComments());
                    break;

                case TokenType.Wend:
                    if (!ImplementsVB60)
                        IsSyntaxError = true;
                    else
                        Statement = ParseEndStatement();
                    break;

                case TokenType.End:
                    Statement = ParseEndStatement();
                    break;

                case TokenType.Return:
                    Statement = ParseExpressionStatement(TreeType.ReturnStatement, true);
                    break;

                case TokenType.RaiseEvent:
                    Statement = ParseRaiseEventStatement();
                    break;

                case TokenType.AddHandler:
                case TokenType.RemoveHandler:
                    Statement = ParseHandlerStatement();
                    break;

                case TokenType.Error:
                    Statement = ParseExpressionStatement(TreeType.ErrorStatement, false);
                    break;

                case TokenType.On:
                    Statement = ParseOnErrorStatement();
                    break;

                case TokenType.Resume:
                    Statement = ParseResumeStatement();
                    break;

                case TokenType.ReDim:
                    Statement = ParseReDimStatement();
                    break;

                case TokenType.Erase:
                    Statement = ParseEraseStatement();
                    break;

                case TokenType.Call:
                    Statement = ParseCallStatement();
                    break;

                case TokenType.IntegerLiteral:
                    if (AtBeginningOfLine)
                        Statement = ParseLabelStatement();
                    else
                        IsSyntaxError = true;
                    break;

                case TokenType.Identifier:
                    Statement = ParseIdentifierStatement();
                    break;

                case TokenType.Period:
                case TokenType.Exclamation:
                case TokenType.Me:
                case TokenType.MyBase:
                case TokenType.MyClass:
                case TokenType.Boolean:
                case TokenType.Byte:
                case TokenType.Short:
                case TokenType.Integer:
                case TokenType.Long:
                case TokenType.Decimal:
                case TokenType.Single:
                case TokenType.Double:
                case TokenType.Date:
                case TokenType.Char:
                case TokenType.String:
                case TokenType.Object:
                case TokenType.DirectCast:
                case TokenType.CType:
                case TokenType.CBool:
                case TokenType.CByte:
                case TokenType.CShort:
                case TokenType.CInt:
                case TokenType.CLng:
                case TokenType.CDec:
                case TokenType.CSng:
                case TokenType.CDbl:
                case TokenType.CDate:
                case TokenType.CChar:
                case TokenType.CStr:
                case TokenType.CVar:
                case TokenType.CObj:
                case TokenType.GetType:
                    Statement = ParseAssignmentOrCallStatement();
                    break;

                case TokenType.Set:
                case TokenType.Let:
                    if (ImplementsVB60)
                        Statement = ParseAssignmentStatement();
                    else
                        IsSyntaxError = true;
                    break;

                case TokenType.Public:
                case TokenType.Private:
                case TokenType.Protected:
                case TokenType.Friend:
                case TokenType.Static:
                case TokenType.Shared:
                case TokenType.Shadows:
                case TokenType.Overloads:
                case TokenType.MustInherit:
                case TokenType.NotInheritable:
                case TokenType.Overrides:
                case TokenType.NotOverridable:
                case TokenType.Overridable:
                case TokenType.MustOverride:
                case TokenType.Partial:
                case TokenType.ReadOnly:
                case TokenType.WriteOnly:
                case TokenType.Dim:
                case TokenType.Const:
                case TokenType.Default:
                case TokenType.WithEvents:
                case TokenType.Widening:
                case TokenType.Narrowing:
                    Statement = ParseLocalDeclarationStatement();
                    break;

                case TokenType.With:
                    Statement = ParseExpressionBlockStatement(TreeType.WithBlockStatement);
                    break;

                case TokenType.SyncLock:
                    Statement = ParseExpressionBlockStatement(TreeType.SyncLockBlockStatement);
                    break;

                case TokenType.Using:
                    Statement = ParseUsingBlockStatement();
                    break;

                case TokenType.While:
                    Statement = ParseExpressionBlockStatement(TreeType.WhileBlockStatement);
                    break;

                case TokenType.Do:
                    Statement = ParseDoBlockStatement();
                    break;

                case TokenType.Loop:
                    Statement = ParseLoopStatement();
                    break;

                case TokenType.For:
                    Statement = ParseForBlockStatement();
                    break;

                case TokenType.Next:
                    Statement = ParseNextStatement();
                    break;

                case TokenType.Throw:
                    Statement = ParseExpressionStatement(TreeType.ThrowStatement, true);
                    break;

                case TokenType.Try:
                    Statement = ParseTryBlockStatement();
                    break;

                case TokenType.Catch:
                    Statement = ParseCatchStatement();
                    break;

                case TokenType.Finally:
                    Statement = new FinallyStatement(SpanFrom(Read()), ParseTrailingComments());
                    break;

                case TokenType.Select:
                    Statement = ParseSelectBlockStatement();
                    break;

                case TokenType.Case:
                    Statement = ParseCaseStatement();
                    break;

                case TokenType.If:
                    Statement = ParseIfBlockStatement();
                    break;

                case TokenType.Else:
                    Statement = new ElseStatement(SpanFrom(Read()), ParseTrailingComments());
                    break;

                case TokenType.ElseIf:
                    Statement = ParseElseIfStatement();
                    break;

                case TokenType.LineTerminator:
                case TokenType.Colon:
                    // An empty statement
                    break;
                
                case TokenType.Comment:
                    Statement = ParseComments();
                    break;

                default:
                    IsSyntaxError = true;
                    break;
            }

            if (IsSyntaxError)
                ReportSyntaxError(SyntaxErrorType.SyntaxError, Start);

            terminator = VerifyEndOfStatement(Statement);

            return Statement;
        }

        private Statement ParseIdentifierStatement()
        {
            Token Start = Peek();
            Statement Statement = null;

            bool IsLabel = false;

            if (AtBeginningOfLine)
            {
                Read();
                Token Current = Peek();
                IsLabel = Current.Type == TokenType.Colon;
                GoBack(Start);
            }

            if (IsLabel)
            {
                Statement = ParseLabelStatement();
            }
            else
            {
                if (Start.AsUnreservedKeyword() == TokenType.Mid)
                    Statement = ParseMidAssignmentStatement();

                if (Statement == null)
                    Statement = ParseAssignmentOrCallStatement();
            }

            return Statement;
        }

        private Statement ParseAssignmentOrCallStatement()
        {
            Statement Result;
            Expression Target = ParseBinaryOperatorExpression(PrecedenceLevel.Power);
            if (ErrorInConstruct)
                ResyncAt(TokenType.Equals);

            // Could be a function call or it could be an assignment
            if (GetAssignmentOperator(Peek().Type) != TreeType.SyntaxError)
                Result = ParseAssignmentStatement(Target);
            else
                Result = ParseCallStatement(Target);

            return Result;
        }

        private StatementCollection ParseStatementBlock(Span blockStartSpan, TreeType blockType, ref List<Comment> Comments)
        {
            Statement empty = null;
            return ParseStatementBlock(blockStartSpan, blockType, ref Comments, ref empty);
        }

        private StatementCollection ParseStatementBlock(Span blockStartSpan, TreeType blockType, ref List<Comment> Comments, ref Statement endStatement)
        {
            List<Statement> Statements = new List<Statement>();
            List<Location> ColonLocations = new List<Location>();
            Token Terminator;
            Token Start;
            Location StatementsEnd;
            bool BlockTerminated = false;

            Debug.Assert(blockType != TreeType.LineIfBlockStatement);
            Comments = ParseTrailingComments();
            Terminator = VerifyEndOfStatement();

            if (Terminator.Type == TokenType.Colon)
            {
                if (blockType == TreeType.SubDeclaration || blockType == TreeType.FunctionDeclaration || blockType == TreeType.ConstructorDeclaration || blockType == TreeType.OperatorDeclaration || blockType == TreeType.GetAccessorDeclaration || blockType == TreeType.SetAccessorDeclaration || blockType == TreeType.AddHandlerAccessorDeclaration || blockType == TreeType.RemoveHandlerAccessorDeclaration || blockType == TreeType.RaiseEventAccessorDeclaration)
                    ReportSyntaxError(SyntaxErrorType.MethodBodyNotAtLineStart, Terminator.Span);

                ColonLocations.Add(Terminator.Span.Start);
            }

            Start = Peek();
            StatementsEnd = Start.Span.Finish;
            endStatement = null;

            PushBlockContext(blockType);

            while (Peek().Type != TokenType.EndOfStream)
            {
                Token PreviousTerminator = Terminator;
                Statement Statement;

                Statement = ParseStatement(ref Terminator);

                if (Statement != null)
                {
                    if (Statement.Type >= TreeType.LoopStatement && Statement.Type <= TreeType.EndBlockStatement)
                    {
                        if (StatementEndsBlock(blockType, Statement))
                        {
                            endStatement = Statement;
                            GoBack(Terminator);
                            BlockTerminated = true;
                            break;
                        }
                        else
                        {
                            bool StatementEndsOuterBlock = false;

                            // If the end statement matches an outer block context, then we want to unwind
                            // up to that level. Otherwise, we want to just give an error and keep going.
                            foreach (TreeType BlockContext in BlockContextStack)
                            {
                                if (StatementEndsBlock(BlockContext, Statement))
                                {
                                    StatementEndsOuterBlock = true;
                                    break;
                                }
                            }

                            if (StatementEndsOuterBlock)
                            {
                                ReportMismatchedEndError(blockType, Statement.Span);
                                // CONSIDER: Can we avoid parsing and re-parsing this statement?
                                GoBack(PreviousTerminator);
                                // We consider the block terminated.
                                BlockTerminated = true;
                                break;
                            }
                            else
                            {
                                ReportMissingBeginStatementError(blockType, Statement);
                            }
                        }
                    }

                    Statements.Add(Statement);
                }

                if (Terminator.Type == TokenType.Colon)
                {
                    ColonLocations.Add(Terminator.Span.Start);
                    StatementsEnd = Terminator.Span.Finish;
                }
                else
                {
                    StatementsEnd = Terminator.Span.Finish;
                }
            }

            if (!BlockTerminated)
                ReportMismatchedEndError(blockType, blockStartSpan);

            PopBlockContext();

            if (Statements.Count == 0 && ColonLocations.Count == 0)
                return null;
            else
                return new StatementCollection(Statements, ColonLocations, new Span(Start.Span.Start, StatementsEnd));
        }

        private StatementCollection ParseLineIfStatementBlock()
        {
            VerifyEndOfCallStatement();

            List<Statement> Statements = new List<Statement>();
            List<Location> ColonLocations = new List<Location>();
            Token Terminator = null;
            Token Start;
            Location StatementsEnd;

            Start = Peek();
            StatementsEnd = Start.Span.Finish;

            PushBlockContext(TreeType.LineIfBlockStatement);

            while (!CanEndStatement(Peek()))
            {
                Statement Statement;

                Statement = ParseStatement(ref Terminator);

                if (Statement != null)
                {
                    if (Statement.Type >= TreeType.LoopStatement && Statement.Type <= TreeType.EndBlockStatement)
                        ReportSyntaxError(SyntaxErrorType.EndInLineIf, Statement.Span);

                    Statements.Add(Statement);
                }

                if (Terminator.Type == TokenType.Colon)
                {
                    ColonLocations.Add(Terminator.Span.Start);
                    StatementsEnd = Terminator.Span.Finish;
                }
                else
                {
                    GoBack(Terminator);
                    break;
                }
            }

            PopBlockContext();
            StatementCollection Result = null;

            if (Statements.Count != 0 || ColonLocations.Count != 0)
            {
                Span span = new Span(Start.Span.Start, StatementsEnd);
                Result = new StatementCollection(Statements, ColonLocations, span);
            }
            
            return Result;
        }

        private Statement ParseComments()
        {
            List<Comment> Comments = new List<Comment>();
            Token LastTerminator;
            Token Start = Peek();

            do
            {
                CommentToken CommentToken = (CommentToken)Scanner.Read();
                Comments.Add(new Comment(CommentToken.Comment, CommentToken.IsREM, CommentToken.Span));
                LastTerminator = Read();
                // Eat the terminator of the comment
            }
            while (Peek().Type == TokenType.Comment);

            GoBack(LastTerminator);

            EmptyStatement statement = new EmptyStatement(SpanFrom(Start), new List<Comment>(Comments));
            return statement;
        }

        #endregion

        #region Modifiers

        private void ValidateModifierList(ModifierCollection modifiers, ModifierTypes validTypes)
        {
            if (modifiers == null)
                return;

            foreach (Modifier Modifier in modifiers)
            {
                if ((validTypes & Modifier.ModifierType) == 0)
                    ReportSyntaxError(SyntaxErrorType.InvalidModifier, Modifier.Span);
            }
        }

        private ModifierCollection ParseDeclarationModifierList()
        {
            List<Modifier> Modifiers = new List<Modifier>();
            Token Start = Peek();
            ModifierTypes ModifierTypes = ModifierTypes.None;
            ModifierTypes FoundTypes = ModifierTypes.None;
            bool found = false;

            while (!found)
            {
                switch (Peek().Type)
                {
                    case TokenType.Public:
                        ModifierTypes = ModifierTypes.Public;
                        break;

                    case TokenType.Private:
                        ModifierTypes = ModifierTypes.Private;
                        break;

                    case TokenType.Protected:
                        ModifierTypes = ModifierTypes.Protected;
                        break;

                    case TokenType.Friend:
                        ModifierTypes = ModifierTypes.Friend;
                        break;

                    case TokenType.Static:
                        ModifierTypes = ModifierTypes.Static;
                        break;

                    case TokenType.Shared:
                        ModifierTypes = ModifierTypes.Shared;
                        break;

                    case TokenType.Shadows:
                        ModifierTypes = ModifierTypes.Shadows;
                        break;

                    case TokenType.Overloads:
                        ModifierTypes = ModifierTypes.Overloads;
                        break;

                    case TokenType.MustInherit:
                        ModifierTypes = ModifierTypes.MustInherit;
                        break;

                    case TokenType.NotInheritable:
                        ModifierTypes = ModifierTypes.NotInheritable;
                        break;

                    case TokenType.Overrides:
                        ModifierTypes = ModifierTypes.Overrides;
                        break;

                    case TokenType.Overridable:
                        ModifierTypes = ModifierTypes.Overridable;
                        break;

                    case TokenType.NotOverridable:
                        ModifierTypes = ModifierTypes.NotOverridable;
                        break;

                    case TokenType.MustOverride:
                        ModifierTypes = ModifierTypes.MustOverride;
                        break;

                    case TokenType.Partial:
                        ModifierTypes = ModifierTypes.Partial;
                        break;

                    case TokenType.ReadOnly:
                        ModifierTypes = ModifierTypes.ReadOnly;
                        break;

                    case TokenType.WriteOnly:
                        ModifierTypes = ModifierTypes.WriteOnly;
                        break;

                    case TokenType.Dim:
                        ModifierTypes = ModifierTypes.Dim;
                        break;

                    case TokenType.Const:
                        ModifierTypes = ModifierTypes.Const;
                        break;

                    case TokenType.Default:
                        ModifierTypes = ModifierTypes.Default;
                        break;

                    case TokenType.WithEvents:
                        ModifierTypes = ModifierTypes.WithEvents;
                        break;

                    case TokenType.Widening:
                        ModifierTypes = ModifierTypes.Widening;
                        break;

                    case TokenType.Narrowing:
                        ModifierTypes = ModifierTypes.Narrowing;
                        break;

                    default:
                        found = true;
                        break;
                }

                if (found)
                    break;

                if ((FoundTypes & ModifierTypes) != 0)
                    ReportSyntaxError(SyntaxErrorType.DuplicateModifier, Peek());
                else
                    FoundTypes = FoundTypes | ModifierTypes;

                Modifiers.Add(new Modifier(ModifierTypes, SpanFrom(Read())));
            }

            if (Modifiers.Count == 0)
                return null;
            else
                return new ModifierCollection(Modifiers, SpanFrom(Start));
        }

        private ModifierCollection ParseParameterModifierList()
        {
            List<Modifier> Modifiers = new List<Modifier>();
            Token Start = Peek();
            ModifierTypes ModifierTypes = ModifierTypes.None;
            ModifierTypes FoundTypes = ModifierTypes.None;
            bool found = false;

            while (!found)
            {
                switch (Peek().Type)
                {
                    case TokenType.ByVal:
                        ModifierTypes = ModifierTypes.ByVal;
                        break;

                    case TokenType.ByRef:
                        ModifierTypes = ModifierTypes.ByRef;
                        break;

                    case TokenType.Optional:
                        ModifierTypes = ModifierTypes.Optional;
                        break;

                    case TokenType.ParamArray:
                        ModifierTypes = ModifierTypes.ParamArray;
                        break;

                    default:
                        found = true;
                        break;
                }

                if (found)
                    break;

                if ((FoundTypes & ModifierTypes) != 0)
                    ReportSyntaxError(SyntaxErrorType.DuplicateModifier, Peek());
                else
                    FoundTypes = FoundTypes | ModifierTypes;

                Modifiers.Add(new Modifier(ModifierTypes, SpanFrom(Read())));
            }

            if (Modifiers.Count == 0)
                return null;
            else
                return new ModifierCollection(Modifiers, SpanFrom(Start));
        }

        #endregion

        #region VariableDeclarators

        private VariableDeclarator ParseVariableDeclarator(ModifierCollection modifiers)
        {
            Token DeclarationStart = Peek();
            List<Location> VariableNamesCommaLocations = new List<Location>();
            List<VariableName> VariableNames = new List<VariableName>();
            Location AsLocation = Location.Empty;
            Location NewLocation = Location.Empty;
            TypeName Type = null;
            ArgumentCollection NewArguments = null;
            Location EqualsLocation = Location.Empty;
            Initializer Initializer = null;
            VariableNameCollection VariableNameCollection = null;

            // Parse the declarators
            do
            {
                VariableName VariableName;

                if (VariableNames.Count > 0)
                    VariableNamesCommaLocations.Add(ReadLocation());
                
                VariableName = ParseVariableName(true);

                if (ErrorInConstruct)
                    ResyncAt(TokenType.As, TokenType.Comma, TokenType.New, TokenType.Equals);
                
                VariableNames.Add(VariableName);
            }
            while (Peek().Type == TokenType.Comma);

            VariableNameCollection = new VariableNameCollection(VariableNames, VariableNamesCommaLocations, SpanFrom(DeclarationStart));

            if (Peek().Type == TokenType.As)
            {
                AsLocation = ReadLocation();

                if (Peek().Type == TokenType.New)
                {
                    NewLocation = ReadLocation();
                    Type = ParseTypeName(false);
                    NewArguments = ParseArguments();
                }
                else
                {
                    Type = ParseTypeName(true);

                    if (ErrorInConstruct)
                        ResyncAt(TokenType.Comma, TokenType.Equals);
                }
            }

            if (Type is ArrayTypeName)
            {
                foreach (VariableName variable in VariableNameCollection)
                {
                    if (variable.ArrayType != null)
                        ReportSyntaxError(SyntaxErrorType.ExpectedEndOfStatement, ((ArrayTypeName)Type).Arguments.Span);
                }
            }

            bool HasInitializer = Peek().Type == TokenType.Equals;
            bool HasNew = NewLocation.IsValid;
            bool IsConst = modifiers != null && modifiers.IsOfType(ModifierTypes.Const);
            bool CanHaveInitializer = (ImplementsVB60 && IsConst) || (!ImplementsVB60 && !HasNew);

            if (HasInitializer && CanHaveInitializer)
            {
                EqualsLocation = ReadLocation();
                Initializer = ParseInitializer();

                if (ErrorInConstruct)
                    ResyncAt(TokenType.Comma);
            }

            return new VariableDeclarator(VariableNameCollection, AsLocation, NewLocation, Type, NewArguments, EqualsLocation, Initializer, SpanFrom(DeclarationStart));
        }

        private VariableDeclaratorCollection ParseVariableDeclarators(ModifierCollection modifiers)
        {
            Token Start = Peek();
            List<VariableDeclarator> VariableDeclarators = new List<VariableDeclarator>();
            List<Location> DeclarationsCommaLocations = new List<Location>();

            // Parse the declarations
            do
            {
                if (VariableDeclarators.Count > 0)
                    DeclarationsCommaLocations.Add(ReadLocation());
            
                VariableDeclarators.Add(ParseVariableDeclarator(modifiers));
            }
            while (Peek().Type == TokenType.Comma);

            return new VariableDeclaratorCollection(VariableDeclarators, DeclarationsCommaLocations, SpanFrom(Start));
        }

        private VariableDeclarator ParseForLoopVariableDeclarator(ref Expression controlExpression)
        {
            Token Start = Peek();
            Location AsLocation = Location.Empty;
            TypeName Type = null;
            VariableName VariableName = null;
            List<VariableName> VariableNames = new List<VariableName>();
            VariableNameCollection VariableNameCollection = null;

            VariableName = ParseVariableName(false);
            VariableNames.Add(VariableName);
            VariableNameCollection = new VariableNameCollection(VariableNames, null, SpanFrom(Start));

            if (ErrorInConstruct)
            {
                // If we see As before a In or Each, then assume that we are still on the Control Variable Declaration. 
                // Otherwise, don't resync and allow the caller to decide how to recover.
                if (PeekAheadFor(TokenType.As, TokenType.In, TokenType.Equals) == TokenType.As)
                {
                    ResyncAt(TokenType.As);
                }
            }

            if (Peek().Type == TokenType.As)
            {
                AsLocation = ReadLocation();
                Type = ParseTypeName(true);
            }

            controlExpression = new SimpleNameExpression(VariableName.Name, VariableName.Span);

            return new VariableDeclarator(VariableNameCollection, AsLocation, Location.Empty, Type, null, Location.Empty, null, SpanFrom(Start));
        }

        #endregion

        #region CaseClauses

        private CaseClause ParseCase()
        {
            Token Start = Peek();

            if (Start.Type == TokenType.Is || IsRelationalOperator(Start.Type))
            {
                Location IsLocation = Location.Empty;
                Token OperatorToken = null;
                BinaryOperatorType Operator = BinaryOperatorType.None;
                Expression Operand = null;

                if (Start.Type == TokenType.Is)
                {
                    IsLocation = ReadLocation();
                }

                if (IsRelationalOperator(Peek().Type))
                {
                    OperatorToken = Read();
                    Operator = GetBinaryOperator(OperatorToken.Type);
                    Operand = ParseExpression();

                    if (ErrorInConstruct)
                    {
                        ResyncAt();
                    }

                    return new ComparisonCaseClause(IsLocation, Operator, OperatorToken.Span.Start, Operand, SpanFrom(Start));
                }
                else
                {
                    ReportSyntaxError(SyntaxErrorType.ExpectedRelationalOperator, Peek());
                    ResyncAt();
                    return null;
                }
            }
            else
            {
                return new RangeCaseClause(ParseExpression(true), SpanFrom(Start));
            }
        }

        #endregion

        #region Attributes

        private AttributeCollection ParseAttributeBlock(AttributeTypes attributeTypesAllowed)
        {
            Token Start = Peek();
            List<AttributeTree> Attributes = new List<AttributeTree>();
            Location RightBracketLocation;
            List<Location> CommaLocations = new List<Location>();

            if (Start.Type != TokenType.LessThan)
            {
                return null;
            }

            Read();

            do
            {
                Token AttributeStart;
                AttributeTypes AttributeTypes = AttributeTypes.Regular;
                Location AttributeTypeLocation = new Location();
                Location ColonLocation = new Location();
                Name Name;
                ArgumentCollection Arguments;

                if (Attributes.Count > 0)
                {
                    CommaLocations.Add(ReadLocation());
                }

                AttributeStart = Peek();

                if (AttributeStart.AsUnreservedKeyword() == TokenType.Assembly)
                {
                    AttributeTypes = AttributeTypes.Assembly;
                    AttributeTypeLocation = ReadLocation();
                    ColonLocation = VerifyExpectedToken(TokenType.Colon);
                }
                else if (AttributeStart.Type == TokenType.Module)
                {
                    AttributeTypes = AttributeTypes.Module;
                    AttributeTypeLocation = ReadLocation();
                    ColonLocation = VerifyExpectedToken(TokenType.Colon);
                }

                if ((AttributeTypes & attributeTypesAllowed) == 0)
                    ReportSyntaxError(SyntaxErrorType.IncorrectAttributeType, AttributeStart);

                Name = ParseName(true);
                Arguments = ParseArguments();

                Attributes.Add(new AttributeTree(AttributeTypes, AttributeTypeLocation, ColonLocation, Name, Arguments, SpanFrom(AttributeStart)));
            }
            while (Peek().Type == TokenType.Comma);

            RightBracketLocation = VerifyExpectedToken(TokenType.GreaterThan);

            return new AttributeCollection(Attributes, CommaLocations, RightBracketLocation, SpanFrom(Start));
        }

        private AttributeBlockCollection ParseAttributes()
        {
            return ParseAttributes(AttributeTypes.Regular);
        }

        private AttributeBlockCollection ParseAttributes(AttributeTypes attributeTypesAllowed)
        {
            Token Start = Peek();
            List<AttributeCollection> AttributeBlocks = new List<AttributeCollection>();

            while (Peek().Type == TokenType.LessThan)
            {
                AttributeBlocks.Add(ParseAttributeBlock(attributeTypesAllowed));
            }

            if (AttributeBlocks.Count == 0)
            {
                return null;
            }
            else
            {
                return new AttributeBlockCollection(AttributeBlocks, SpanFrom(Start));
            }
        }

        #endregion

        #region Declaration statements

        private NameCollection ParseNameList()
        {
            return ParseNameList(false);
        }

        private NameCollection ParseNameList(bool allowLeadingMeOrMyBase)
        {
            Token Start = Read();
            List<Location> CommaLocations = new List<Location>();
            List<Name> Names = new List<Name>();

            do
            {
                if (Names.Count > 0)
                    CommaLocations.Add(ReadLocation());
                
                Names.Add(ParseNameListName(allowLeadingMeOrMyBase));

                if (ErrorInConstruct)
                    ResyncAt(TokenType.Comma);
            }
            while (Peek().Type == TokenType.Comma);

            return new NameCollection(Names, CommaLocations, SpanFrom(Start));
        }

        private Declaration ParsePropertyDeclaration(Location startLocation, AttributeBlockCollection attributes, ModifierCollection modifiers)
        {
            const ModifierTypes RegularValidModifiers = 
                ModifierTypes.AccessModifiers | 
                ModifierTypes.Shadows | 
                ModifierTypes.Shared | 
                ModifierTypes.Overridable | 
                ModifierTypes.NotOverridable | 
                ModifierTypes.MustOverride | 
                ModifierTypes.Overrides | 
                ModifierTypes.Overloads | 
                ModifierTypes.Default | 
                ModifierTypes.ReadOnly | 
                ModifierTypes.WriteOnly;

            const ModifierTypes InterfaceValidModifiers = 
                ModifierTypes.Shadows | 
                ModifierTypes.Overloads | 
                ModifierTypes.Default | 
                ModifierTypes.ReadOnly | 
                ModifierTypes.WriteOnly;

            Location PropertyLocation = Location.Empty;
            SimpleName Name = null;
            ParameterCollection Parameters = null;
            Location AsLocation = Location.Empty;
            TypeName ReturnType = null;
            AttributeBlockCollection ReturnTypeAttributes = null;
            NameCollection ImplementsList = null;
            
            StatementCollection Statements = null;
            DeclarationCollection Accessors = null;
            GetAccessorDeclaration getAcessor = null;
            SetAccessorDeclaration setAcessor = null;

            EndBlockDeclaration EndBlockDeclaration = null;
            Statement EndBlockStatement = null;
            List<Comment> Comments = null;
            TypeParameterCollection TypeParameters = null;

            bool InInterface = CurrentBlockContextType() == TreeType.InterfaceDeclaration;
            Span blockStartSpan;
            ModifierTypes ValidModifiers = InInterface ? InterfaceValidModifiers : RegularValidModifiers;
            Token AccessorToken = null;

            ValidateModifierList(modifiers, ValidModifiers);
            PropertyLocation = ReadLocation();
            
            if (ImplementsVB60)
            {
                VerifyExpectedToken(true, TokenType.Get, TokenType.Let, TokenType.Set);
                if (ErrorInConstruct)
                    ResyncAt(TokenType.Property);
                else
                    AccessorToken = Read();

                if (AccessorToken != null && IdentifierToken.IsIdentifier(AccessorToken.Type))
                {
                    IdentifierToken Identifier = (IdentifierToken)AccessorToken;
                    AccessorToken = new IdentifierToken(Identifier.AsUnreservedKeyword(), Identifier.AsUnreservedKeyword(), Identifier.Identifier, Identifier.Escaped, Identifier.HasSpaceAfter, Identifier.TypeCharacter, Identifier.Span);
                }
            }

            Name = ParseSimpleName(false);
            if (ErrorInConstruct)
                ResyncAt(TokenType.LeftParenthesis, TokenType.As);

            TypeParameters = ParseTypeParameters();
            if (ErrorInConstruct)
                ResyncAt(TokenType.LeftParenthesis, TokenType.As);
            if (TypeParameters != null && TypeParameters.Count > 0)
                ReportSyntaxError(SyntaxErrorType.PropertiesCantBeGeneric, TypeParameters.Span);

            Parameters = ParseParameters();

            // Let and Set must have arguments
            if (ImplementsVB60 && 
                AccessorToken != null && (AccessorToken.Type == TokenType.Let || AccessorToken.Type == TokenType.Set) && 
                Parameters != null && Parameters.Count == 0)
            {
                // Argument required for Property Let or Property Set
                ReportSyntaxError(SyntaxErrorType.PropertyArgumentRequired, Peek());
            }

            if (Peek().Type == TokenType.As)
            {
                AsLocation = ReadLocation();
                ReturnTypeAttributes = ParseAttributes();
                ReturnType = ParseTypeName(true);

                if (ErrorInConstruct)
                    ResyncAt(TokenType.Implements);
            }

            if (ImplementsVB60)
            {
                blockStartSpan = SpanFrom(startLocation);
                Statements = ParseStatementBlock(blockStartSpan, TreeType.PropertyDeclaration, ref Comments, ref EndBlockStatement);

                if (EndBlockDeclaration == null)
                    EndBlockDeclaration = new EndBlockDeclaration(BlockType.Property, Location.Empty, Span.Empty, null);

                List<Declaration> declatarions = new List<Declaration>();

                switch (AccessorToken.Type)
                {
                    case TokenType.Get:
                        getAcessor = new GetAccessorDeclaration(null, null, Location.Empty, Statements, null, Span.Empty, null);
                        declatarions.Add(getAcessor);
                        break;
                    case TokenType.Let:
                    case TokenType.Set:
                        setAcessor = new SetAccessorDeclaration(null, null, Location.Empty, Parameters, Statements, null, Span.Empty, null);
                        declatarions.Add(setAcessor);
                        break;
                }
            }
            else if (!InInterface)
            {
                if (Peek().Type == TokenType.Implements)
                    ImplementsList = ParseNameList();

                if (modifiers == null || !modifiers.IsOfType(ModifierTypes.MustOverride))
                {
                    blockStartSpan = SpanFrom(startLocation);
                    Accessors = ParseDeclarationBlock(blockStartSpan, TreeType.PropertyDeclaration, ref Comments, ref EndBlockDeclaration);
                }

                if (Accessors != null)
                {
                    foreach (Declaration item in Accessors)
                    {
                        if (item is GetAccessorDeclaration)
                            getAcessor = (GetAccessorDeclaration)item;
                        else if (item is SetAccessorDeclaration)
                            setAcessor = (SetAccessorDeclaration)item;
                    }
                }
            }

            if (Comments == null)
                Comments = ParseTrailingComments();

            blockStartSpan = SpanFrom(startLocation);

            return new PropertyDeclaration(
                attributes, 
                modifiers, 
                PropertyLocation, 
                Name, 
                Parameters, 
                AsLocation, 
                ReturnTypeAttributes, 
                ReturnType, 
                ImplementsList, 
                getAcessor,
                setAcessor,
                EndBlockDeclaration, 
                blockStartSpan, 
                Comments
            );
        }

        private Declaration ParseExternalDeclaration(Location startLocation, AttributeBlockCollection attributes, ModifierCollection modifiers)
        {
            const ModifierTypes ValidModifiers = ModifierTypes.AccessModifiers | ModifierTypes.Shadows | ModifierTypes.Overloads;
            Location DeclareLocation = Location.Empty;
            Location CharsetLocation = Location.Empty;
            Charset Charset = Charset.Auto;
            TreeType MethodType = TreeType.SyntaxError;
            Location SubOrFunctionLocation = Location.Empty;
            SimpleName Name = null;
            Location LibLocation = Location.Empty;
            StringLiteralExpression LibLiteral = null;
            Location AliasLocation = Location.Empty;
            StringLiteralExpression AliasLiteral = null;
            ParameterCollection Parameters = null;
            Location AsLocation = Location.Empty;
            TypeName ReturnType = null;
            AttributeBlockCollection ReturnTypeAttributes = null;

            ValidateModifierList(modifiers, ValidModifiers);

            DeclareLocation = ReadLocation();
            
            Token CharsetToken = Peek();
            TokenType CharsetTokenType = CharsetToken.AsUnreservedKeyword();
            switch (CharsetTokenType)
            {
                case TokenType.Ansi:
                    Charset = Charset.Ansi;
                    CharsetLocation = ReadLocation();
                    break;

                case TokenType.Unicode:
                    Charset = Charset.Unicode;
                    CharsetLocation = ReadLocation();
                    break;

                case TokenType.Auto:
                    Charset = Charset.Auto;
                    CharsetLocation = ReadLocation();
                    break;
            }

            if (ImplementsVB60 && CharsetLocation.IsValid)
                ReportSyntaxError(SyntaxErrorType.ExpectedSubOrFunction, CharsetToken.Span);

            if (Peek().Type == TokenType.Sub)
            {
                MethodType = TreeType.ExternalSubDeclaration;
                SubOrFunctionLocation = ReadLocation();
            }
            else if (Peek().Type == TokenType.Function)
            {
                MethodType = TreeType.ExternalFunctionDeclaration;
                SubOrFunctionLocation = ReadLocation();
            }
            else
            {
                ReportSyntaxError(SyntaxErrorType.ExpectedSubOrFunction, Peek());
            }

            Name = ParseSimpleName(false);

            if (ErrorInConstruct)
            {
                ResyncAt(TokenType.Lib, TokenType.LeftParenthesis);
            }

            if (Peek().Type == TokenType.Lib)
            {
                LibLocation = ReadLocation();

                if (Peek().Type == TokenType.StringLiteral)
                {
                    StringLiteralToken Literal = (StringLiteralToken)Read();
                    LibLiteral = new StringLiteralExpression(Literal.Literal, Literal.Span);
                }
                else
                {
                    ReportSyntaxError(SyntaxErrorType.ExpectedStringLiteral, Peek());
                    ResyncAt(TokenType.Alias, TokenType.LeftParenthesis);
                }
            }
            else
            {
                ReportSyntaxError(SyntaxErrorType.ExpectedLib, Peek());
            }

            if (Peek().Type == TokenType.Alias)
            {
                AliasLocation = ReadLocation();

                if (Peek().Type == TokenType.StringLiteral)
                {
                    StringLiteralToken Literal = (StringLiteralToken)Read();
                    AliasLiteral = new StringLiteralExpression(Literal.Literal, Literal.Span);
                }
                else
                {
                    ReportSyntaxError(SyntaxErrorType.ExpectedStringLiteral, Peek());
                    ResyncAt(TokenType.LeftParenthesis);
                }
            }

            Parameters = ParseParameters();

            if (MethodType == TreeType.ExternalFunctionDeclaration)
            {
                if (Peek().Type == TokenType.As)
                {
                    AsLocation = ReadLocation();
                    ReturnTypeAttributes = ParseAttributes();
                    ReturnType = ParseTypeName(true);

                    if (ErrorInConstruct)
                    {
                        ResyncAt();
                    }
                }

                return new ExternalFunctionDeclaration(attributes, modifiers, DeclareLocation, CharsetLocation, Charset, SubOrFunctionLocation, Name, LibLocation, LibLiteral, AliasLocation,
                AliasLiteral, Parameters, AsLocation, ReturnTypeAttributes, ReturnType, SpanFrom(startLocation), ParseTrailingComments());
            }
            else if (MethodType == TreeType.ExternalSubDeclaration)
            {
                return new ExternalSubDeclaration(attributes, modifiers, DeclareLocation, CharsetLocation, Charset, SubOrFunctionLocation, Name, LibLocation, LibLiteral, AliasLocation,
                AliasLiteral, Parameters, SpanFrom(startLocation), ParseTrailingComments());
            }
            else
            {
                return Declaration.GetBadDeclaration(SpanFrom(startLocation), ParseTrailingComments());
            }
        }

        private Declaration ParseMethodDeclaration(Location startLocation, AttributeBlockCollection attributes, ModifierCollection modifiers)
        {
            const ModifierTypes ValidMethodModifiers = ModifierTypes.AccessModifiers | ModifierTypes.Shadows | ModifierTypes.Shared | ModifierTypes.Overridable | ModifierTypes.NotOverridable | ModifierTypes.MustOverride | ModifierTypes.Overrides | ModifierTypes.Overloads;
            const ModifierTypes ValidConstructorModifiers = ModifierTypes.AccessModifiers | ModifierTypes.Shared;
            const ModifierTypes ValidInterfaceModifiers = ModifierTypes.Shadows | ModifierTypes.Overloads;

            TreeType MethodType;
            Location SubOrFunctionLocation = Location.Empty;
            SimpleName Name = null;
            ParameterCollection Parameters = null;
            Location AsLocation = Location.Empty;
            TypeName ReturnType = null;
            AttributeBlockCollection ReturnTypeAttributes = null;
            NameCollection ImplementsList = null;
            NameCollection HandlesList = null;
            bool AllowKeywordsForName = false;
            ModifierTypes ValidModifiers = ValidMethodModifiers;
            StatementCollection Statements = null;
            Statement EndStatement = null;
            EndBlockDeclaration EndDeclaration = null;
            bool InInterface = CurrentBlockContextType() == TreeType.InterfaceDeclaration;
            List<Comment> Comments = null;
            TypeParameterCollection TypeParameters = null;

            if (!AtBeginningOfLine)
                ReportSyntaxError(SyntaxErrorType.MethodMustBeFirstStatementOnLine, Peek());

            if (Peek().Type == TokenType.Sub)
            {
                SubOrFunctionLocation = ReadLocation();

                if (Peek().Type == TokenType.New)
                {
                    MethodType = TreeType.ConstructorDeclaration;
                    AllowKeywordsForName = true;
                    ValidModifiers = ValidConstructorModifiers;
                }
                else
                {
                    MethodType = TreeType.SubDeclaration;
                }
            }
            else
            {
                SubOrFunctionLocation = ReadLocation();
                MethodType = TreeType.FunctionDeclaration;
            }

            if (InInterface)
            {
                ValidModifiers = ValidInterfaceModifiers;
            }

            ValidateModifierList(modifiers, ValidModifiers);
            Name = ParseSimpleName(AllowKeywordsForName);

            if (ImplementsVB60 && string.Compare(Name.Name, "Class_Initialize", true) == 0)
                MethodType = TreeType.ConstructorDeclaration;

            if (ErrorInConstruct)
                ResyncAt(TokenType.LeftParenthesis, TokenType.As);

            TypeParameters = ParseTypeParameters();

            if (ErrorInConstruct)
                ResyncAt(TokenType.LeftParenthesis, TokenType.As);

            if (MethodType == TreeType.ConstructorDeclaration && TypeParameters != null && TypeParameters.Count > 0)
                ReportSyntaxError(SyntaxErrorType.ConstructorsCantBeGeneric, TypeParameters.Span);


            Parameters = ParseParameters();

            if (MethodType == TreeType.FunctionDeclaration && Peek().Type == TokenType.As)
            {
                AsLocation = ReadLocation();
                ReturnTypeAttributes = ParseAttributes();
                ReturnType = ParseTypeName(true);

                if (ErrorInConstruct)
                {
                    ResyncAt(TokenType.Implements, TokenType.Handles);
                }
            }

            if (InInterface)
            {
                Comments = ParseTrailingComments();
            }
            else
            {
                if (Peek().Type == TokenType.Implements)
                    ImplementsList = ParseNameList();
                else if (Peek().Type == TokenType.Handles)
                    HandlesList = ParseNameList(true);
                
                if (modifiers == null || !modifiers.IsOfType(ModifierTypes.MustOverride))
                    Statements = ParseStatementBlock(SpanFrom(startLocation), MethodType, ref Comments, ref EndStatement);
                else
                    Comments = ParseTrailingComments();
                
                if (EndStatement != null)
                    EndDeclaration = new EndBlockDeclaration((EndBlockStatement)EndStatement);
            }

            if (MethodType == TreeType.SubDeclaration)
            {
                return new SubDeclaration(attributes, modifiers, SubOrFunctionLocation, Name, TypeParameters, Parameters, ImplementsList, HandlesList, Statements, EndDeclaration,
                SpanFrom(startLocation), Comments);
            }
            else if (MethodType == TreeType.FunctionDeclaration)
            {
                return new FunctionDeclaration(attributes, modifiers, SubOrFunctionLocation, Name, TypeParameters, Parameters, AsLocation, ReturnTypeAttributes, ReturnType, ImplementsList,
                HandlesList, Statements, EndDeclaration, SpanFrom(startLocation), Comments);
            }
            else
            {
                return new ConstructorDeclaration(attributes, modifiers, SubOrFunctionLocation, Name, Parameters, Statements, EndDeclaration, SpanFrom(startLocation), Comments);
            }
        }


        private Declaration ParseOperatorDeclaration(Location startLocation, AttributeBlockCollection attributes, ModifierCollection modifiers)
        {
            const ModifierTypes ValidOperatorModifiers = ModifierTypes.Shared | ModifierTypes.Public | ModifierTypes.Shadows | ModifierTypes.Overloads | ModifierTypes.Widening | ModifierTypes.Narrowing;

            Location KeywordLocation;
            Token OperatorToken = null;
            ParameterCollection Parameters = null;
            Location AsLocation = Location.Empty;
            TypeName ReturnType = null;
            AttributeBlockCollection ReturnTypeAttributes = null;
            ModifierTypes ValidModifiers = ValidOperatorModifiers;
            StatementCollection Statements = null;
            Statement EndStatement = null;
            EndBlockDeclaration EndDeclaration = null;
            List<Comment> Comments = null;
            TypeParameterCollection TypeParameters = null;

            if (!AtBeginningOfLine)
                ReportSyntaxError(SyntaxErrorType.MethodMustBeFirstStatementOnLine, Peek());

            KeywordLocation = ReadLocation();
            ValidateModifierList(modifiers, ValidModifiers);

            if (IsOverloadableOperator(Peek()))
            {
                OperatorToken = Read();
            }
            else
            {
                ReportSyntaxError(SyntaxErrorType.InvalidOperator, Peek());
                ResyncAt(TokenType.LeftParenthesis, TokenType.As);
            }

            TypeParameters = ParseTypeParameters();

            if (ErrorInConstruct)
            {
                ResyncAt(TokenType.LeftParenthesis, TokenType.As);
            }

            if (TypeParameters != null && TypeParameters.Count > 0)
            {
                ReportSyntaxError(SyntaxErrorType.OperatorsCantBeGeneric, TypeParameters.Span);
            }

            Parameters = ParseParameters();

            if (Peek().Type == TokenType.As)
            {
                AsLocation = ReadLocation();
                ReturnTypeAttributes = ParseAttributes();
                ReturnType = ParseTypeName(true);

                if (ErrorInConstruct)
                {
                    ResyncAt();
                }
            }

            Statements = ParseStatementBlock(SpanFrom(startLocation), TreeType.OperatorDeclaration, ref Comments, ref EndStatement);
            Comments = ParseTrailingComments();

            if (EndStatement != null)
            {
                EndDeclaration = new EndBlockDeclaration((EndBlockStatement)EndStatement);
            }

            return new OperatorDeclaration(attributes, modifiers, KeywordLocation, OperatorToken, Parameters, AsLocation, ReturnTypeAttributes, ReturnType, Statements, EndDeclaration,
            SpanFrom(startLocation), Comments);
        }

        private Declaration ParseAccessorDeclaration(Location startLocation, AttributeBlockCollection attributes, ModifierCollection modifiers)
        {
            TreeType AccessorType;
            Location GetOrSetLocation;
            ParameterCollection Parameters = null;
            StatementCollection Statements;
            Statement EndStatement = null;
            EndBlockDeclaration EndDeclaration = null;
            List<Comment> Comments = null;
            ModifierTypes ValidModifiers = ModifierTypes.None;

            if (ImplementsVB80)
                ValidModifiers = ValidModifiers | ModifierTypes.AccessModifiers;

            if (!AtBeginningOfLine)
                ReportSyntaxError(SyntaxErrorType.MethodMustBeFirstStatementOnLine, Peek());

            ValidateModifierList(modifiers, ValidModifiers);

            if (Peek().Type == TokenType.Get)
                AccessorType = TreeType.GetAccessorDeclaration;
            else
                AccessorType = TreeType.SetAccessorDeclaration;

            GetOrSetLocation = ReadLocation();

            if (AccessorType == TreeType.SetAccessorDeclaration)
                Parameters = ParseParameters();

            Statements = ParseStatementBlock(SpanFrom(startLocation), AccessorType, ref Comments, ref EndStatement);

            if (EndStatement != null)
                EndDeclaration = new EndBlockDeclaration((EndBlockStatement)EndStatement);

            if (AccessorType == TreeType.GetAccessorDeclaration)
                return new GetAccessorDeclaration(attributes, modifiers, GetOrSetLocation, Statements, EndDeclaration, SpanFrom(startLocation), Comments);
            else
                return new SetAccessorDeclaration(attributes, modifiers, GetOrSetLocation, Parameters, Statements, EndDeclaration, SpanFrom(startLocation), Comments);
        }

        private Declaration ParseCustomEventDeclaration(Location startLocation, AttributeBlockCollection attributes, ModifierCollection modifiers)
        {
            const ModifierTypes ValidModifiers = ModifierTypes.AccessModifiers | ModifierTypes.Shadows | ModifierTypes.Shared;

            Location CustomLocation;
            Location EventLocation;
            SimpleName Name;
            Location AsLocation;
            TypeName EventType;
            NameCollection ImplementsList = null;
            DeclarationCollection Accessors = null;
            EndBlockDeclaration EndBlockDeclaration = null;
            List<Comment> Comments = null;

            ValidateModifierList(modifiers, ValidModifiers);
            CustomLocation = ReadLocation();
            Debug.Assert(Peek().Type == TokenType.Event);
            EventLocation = ReadLocation();

            Name = ParseSimpleName(false);

            if (ErrorInConstruct)
            {
                ResyncAt(TokenType.As);
            }

            AsLocation = VerifyExpectedToken(TokenType.As);
            EventType = ParseTypeName(true);

            if (ErrorInConstruct)
            {
                ResyncAt(TokenType.Implements);
            }

            if (Peek().Type == TokenType.Implements)
            {
                ImplementsList = ParseNameList();
            }

            Accessors = ParseDeclarationBlock(SpanFrom(startLocation), TreeType.CustomEventDeclaration, ref Comments, ref EndBlockDeclaration);

            return new CustomEventDeclaration(attributes, modifiers, CustomLocation, EventLocation, Name, AsLocation, EventType, ImplementsList, Accessors, EndBlockDeclaration,
            SpanFrom(startLocation), Comments);
        }

        private Declaration ParseEventAccessorDeclaration(Location startLocation, AttributeBlockCollection attributes, ModifierCollection modifiers)
        {
            const ModifierTypes ValidModifiers = ModifierTypes.None;
            TreeType AccessorType;
            Location AccessorTypeLocation;
            ParameterCollection Parameters = null;
            StatementCollection Statements;
            Statement EndStatement = null;
            EndBlockDeclaration EndDeclaration = null;
            List<Comment> Comments = null;

            if (!AtBeginningOfLine)
            {
                ReportSyntaxError(SyntaxErrorType.MethodMustBeFirstStatementOnLine, Peek());
            }

            ValidateModifierList(modifiers, ValidModifiers);

            if (Peek().Type == TokenType.AddHandler)
            {
                AccessorType = TreeType.AddHandlerAccessorDeclaration;
            }
            else if (Peek().Type == TokenType.RemoveHandler)
            {
                AccessorType = TreeType.RemoveHandlerAccessorDeclaration;
            }
            else
            {
                AccessorType = TreeType.RaiseEventAccessorDeclaration;
            }
            AccessorTypeLocation = ReadLocation();

            Parameters = ParseParameters();
            Statements = ParseStatementBlock(SpanFrom(startLocation), AccessorType, ref Comments, ref EndStatement);

            if (EndStatement != null)
            {
                EndDeclaration = new EndBlockDeclaration((EndBlockStatement)EndStatement);
            }

            if (AccessorType == TreeType.AddHandlerAccessorDeclaration)
            {
                return new AddHandlerAccessorDeclaration(attributes, AccessorTypeLocation, Parameters, Statements, EndDeclaration, SpanFrom(startLocation), Comments);
            }
            else if (AccessorType == TreeType.RemoveHandlerAccessorDeclaration)
            {
                return new RemoveHandlerAccessorDeclaration(attributes, AccessorTypeLocation, Parameters, Statements, EndDeclaration, SpanFrom(startLocation), Comments);
            }
            else
            {
                return new RaiseEventAccessorDeclaration(attributes, AccessorTypeLocation, Parameters, Statements, EndDeclaration, SpanFrom(startLocation), Comments);
            }
        }

        private Declaration ParseEventDeclaration(Location startLocation, AttributeBlockCollection attributes, ModifierCollection modifiers)
        {
            const ModifierTypes RegularValidModifiers = ModifierTypes.AccessModifiers | ModifierTypes.Shadows | ModifierTypes.Shared;
            const ModifierTypes InterfaceValidModifiers = ModifierTypes.Shadows;

            Location EventLocation = Location.Empty;
            SimpleName Name = null; ;
            Location AsLocation = Location.Empty;
            TypeName EventType = null;
            ParameterCollection Parameters = null;
            NameCollection ImplementsList = null;
            bool InInterface = CurrentBlockContextType() == TreeType.InterfaceDeclaration;
            ModifierTypes ValidModifiers;

            if (InInterface)
            {
                ValidModifiers = InterfaceValidModifiers;
            }
            else
            {
                ValidModifiers = RegularValidModifiers;
            }

            ValidateModifierList(modifiers, ValidModifiers);

            EventLocation = ReadLocation();
            Name = ParseSimpleName(false);

            if (ErrorInConstruct)
            {
                ResyncAt(TokenType.As, TokenType.LeftParenthesis, TokenType.Implements);
            }

            if (Peek().Type == TokenType.As)
            {
                AsLocation = ReadLocation();
                EventType = ParseTypeName(false);

                if (ErrorInConstruct)
                {
                    ResyncAt(TokenType.Implements);
                }
            }
            else
            {
                Parameters = ParseParameters();

                // Give a good error if they attempt to do a return type
                if (Peek().Type == TokenType.As)
                {
                    Token ErrorStart = Peek();

                    ResyncAt(TokenType.Implements);
                    ReportSyntaxError(SyntaxErrorType.EventsCantBeFunctions, ErrorStart, Peek());
                }
            }

            if (Peek().Type == TokenType.Implements)
            {
                ImplementsList = ParseNameList();
            }

            return new EventDeclaration(attributes, modifiers, EventLocation, Name, Parameters, AsLocation, null, EventType, ImplementsList, SpanFrom(startLocation), ParseTrailingComments());
        }

        private Declaration ParseVariableListDeclaration(Location startLocation, AttributeBlockCollection attributes, ModifierCollection modifiers)
        {
            ModifierTypes ValidModifiers;

            if (modifiers != null && modifiers.IsOfType(ModifierTypes.Const))
                ValidModifiers = ModifierTypes.Const | ModifierTypes.AccessModifiers | ModifierTypes.Shadows;
            else
                ValidModifiers = ModifierTypes.Dim | ModifierTypes.AccessModifiers | ModifierTypes.Shadows | ModifierTypes.Shared | ModifierTypes.ReadOnly | ModifierTypes.WithEvents;
            
            ValidateModifierList(modifiers, ValidModifiers);

            Token Current = Peek();
            bool IsStructureDeclaration = CurrentBlockContextType() == TreeType.StructureDeclaration;

            if (ImplementsVB60 && IsStructureDeclaration)
            {
                // Members of Type Block Declaration (VB6) doesn't have modifiers
                if (modifiers != null)
                    ReportSyntaxError(SyntaxErrorType.InvalidInsideTypeBlock, Current);
            }
            else if (modifiers == null)
            {
                ReportSyntaxError(SyntaxErrorType.ExpectedModifier, Current);
            }

            VariableDeclaratorCollection declarations = ParseVariableDeclarators(modifiers);
            
            bool IsMultipleDeclaration = declarations != null && declarations.Count > 1;
            // Members of Type Block Declaration (VB6) doesn't have multiple declarations
            if (ImplementsVB60 && IsStructureDeclaration && IsMultipleDeclaration)
            {
                Location errorStart = declarations.CommaLocations[0];
                Location errorFinish = declarations.Span.Finish;
                Span errorSpan = new Span(errorStart, errorFinish);
                ReportSyntaxError(SyntaxErrorType.InvalidInsideTypeBlock, errorSpan);
            }

            IList<Comment> comments = ParseTrailingComments();
            Span start = SpanFrom(startLocation);

            return new VariableListDeclaration(attributes, modifiers, declarations, start, comments);
        }

        private Declaration ParseEndDeclaration()
        {
            Token Start = Read();
            BlockType EndType = GetBlockType(Peek().Type);

            switch (EndType)
            {
                case BlockType.Sub:
                    if (!AtBeginningOfLine)
                        ReportSyntaxError(SyntaxErrorType.EndSubNotAtLineStart, SpanFrom(Start));
                    
                    break;

                case BlockType.Function:
                    if (!AtBeginningOfLine)
                        ReportSyntaxError(SyntaxErrorType.EndFunctionNotAtLineStart, SpanFrom(Start));
                    
                    break;

                case BlockType.Operator:
                    if (!AtBeginningOfLine)
                        ReportSyntaxError(SyntaxErrorType.EndOperatorNotAtLineStart, SpanFrom(Start));
                    
                    break;

                case BlockType.Get:
                    if (!AtBeginningOfLine)
                        ReportSyntaxError(SyntaxErrorType.EndGetNotAtLineStart, SpanFrom(Start));
                    
                    break;

                case BlockType.Set:
                    if (!AtBeginningOfLine)
                        ReportSyntaxError(SyntaxErrorType.EndSetNotAtLineStart, SpanFrom(Start));
                    
                    break;

                case BlockType.AddHandler:
                    if (!AtBeginningOfLine)
                        ReportSyntaxError(SyntaxErrorType.EndAddHandlerNotAtLineStart, SpanFrom(Start));
                    
                    break;

                case BlockType.RemoveHandler:
                    if (!AtBeginningOfLine)
                        ReportSyntaxError(SyntaxErrorType.EndRemoveHandlerNotAtLineStart, SpanFrom(Start));
                    
                    break;

                case BlockType.RaiseEvent:
                    if (!AtBeginningOfLine)
                        ReportSyntaxError(SyntaxErrorType.EndRaiseEventNotAtLineStart, SpanFrom(Start));
                    
                    break;

                case BlockType.None:
                    ReportSyntaxError(SyntaxErrorType.UnrecognizedEnd, Peek());
                    return Declaration.GetBadDeclaration(SpanFrom(Start), ParseTrailingComments());
            }

            return new EndBlockDeclaration(EndType, ReadLocation(), SpanFrom(Start), ParseTrailingComments());
        }

        private Declaration ParseTypeDeclaration(Location startLocation, AttributeBlockCollection attributes, ModifierCollection modifiers, TreeType blockType)
        {
            Span startSpan;
            ModifierTypes ValidModifiers;
            Location KeywordLocation;
            SimpleName Name;
            DeclarationCollection Members;
            EndBlockDeclaration EndBlockDeclaration = null;
            List<Comment> Comments = null;
            TypeParameterCollection TypeParameters = null;

            if (blockType == TreeType.ModuleDeclaration)
            {
                ValidModifiers = ModifierTypes.AccessModifiers;
            }
            else
            {
                ValidModifiers = ModifierTypes.AccessModifiers | ModifierTypes.Shadows;

                if (blockType == TreeType.ClassDeclaration)
                    ValidModifiers = ValidModifiers | ModifierTypes.MustInherit | ModifierTypes.NotInheritable;

                if (blockType == TreeType.ClassDeclaration || blockType == TreeType.StructureDeclaration)
                    ValidModifiers = ValidModifiers | ModifierTypes.Partial;
            }

            ValidateModifierList(modifiers, ValidModifiers);

            KeywordLocation = ReadLocation();

            Name = ParseSimpleName(false);

            if (ErrorInConstruct)
                ResyncAt();

            TypeParameters = ParseTypeParameters();

            if (ErrorInConstruct)
                ResyncAt();

            if (blockType == TreeType.ModuleDeclaration && TypeParameters != null && TypeParameters.Count > 0)
                ReportSyntaxError(SyntaxErrorType.ModulesCantBeGeneric, TypeParameters.Span);

            startSpan = SpanFrom(startLocation);
            Members = ParseDeclarationBlock(startSpan, blockType, ref Comments, ref EndBlockDeclaration);
            
            switch (blockType)
            {
                case TreeType.ClassDeclaration:
                    return new ClassDeclaration(attributes, modifiers, KeywordLocation, Name, TypeParameters, Members, EndBlockDeclaration, SpanFrom(startLocation), Comments);

                case TreeType.ModuleDeclaration:
                    return new ModuleDeclaration(attributes, modifiers, KeywordLocation, Name, Members, EndBlockDeclaration, SpanFrom(startLocation), Comments);

                case TreeType.InterfaceDeclaration:
                    return new InterfaceDeclaration(attributes, modifiers, KeywordLocation, Name, TypeParameters, Members, EndBlockDeclaration, SpanFrom(startLocation), Comments);

                case TreeType.StructureDeclaration:
                    return new StructureDeclaration(attributes, modifiers, KeywordLocation, Name, TypeParameters, Members, EndBlockDeclaration, SpanFrom(startLocation), Comments);

                default:
                    Debug.Fail("unexpected!");
                    return null;
            }
        }

        private Declaration ParseEnumDeclaration(Location startLocation, AttributeBlockCollection attributes, ModifierCollection modifiers)
        {
            const ModifierTypes ValidModifiers = ModifierTypes.AccessModifiers | ModifierTypes.Shadows;
            Location KeywordLocation = Location.Empty;
            SimpleName Name = null;
            Location AsLocation = Location.Empty;
            TypeName Type = null;
            DeclarationCollection Members = null;
            EndBlockDeclaration EndBlockDeclaration = null;
            List<Comment> Comments = null;

            ValidateModifierList(modifiers, ValidModifiers);

            KeywordLocation = ReadLocation();

            Name = ParseSimpleName(false);

            if (ErrorInConstruct)
            {
                ResyncAt(TokenType.As);
            }

            if (Peek().Type == TokenType.As)
            {
                AsLocation = ReadLocation();
                Type = ParseTypeName(false);

                if (ImplementsVB60)
                {
                    Span errorSpan = new Span(AsLocation, AsLocation);
                    if (Type != null && Type.Span != null)
                        errorSpan = new Span(AsLocation, Type.Span.Finish);

                    ReportSyntaxError(SyntaxErrorType.ExpectedEndOfStatement, errorSpan);
                    Type = null;
                }

                if (ErrorInConstruct)
                    ResyncAt();
            }

            Members = ParseDeclarationBlock(SpanFrom(startLocation), TreeType.EnumDeclaration, ref Comments, ref EndBlockDeclaration);

            if (Members == null || Members.Count == 0)
                ReportSyntaxError(SyntaxErrorType.EmptyEnum, Name.Span);

            return new EnumDeclaration(attributes, modifiers, KeywordLocation, Name, AsLocation, Type, Members, EndBlockDeclaration, SpanFrom(startLocation), Comments
            );
        }

        private Declaration ParseDelegateDeclaration(Location startLocation, AttributeBlockCollection attributes, ModifierCollection modifiers)
        {
            const ModifierTypes ValidModifiers = ModifierTypes.AccessModifiers | ModifierTypes.Shadows | ModifierTypes.Shared;
            Location DelegateLocation = Location.Empty;
            TreeType MethodType = TreeType.SyntaxError;
            Location SubOrFunctionLocation = Location.Empty;
            SimpleName Name = null;
            ParameterCollection Parameters = null;
            Location AsLocation = Location.Empty;
            TypeName ReturnType = null;
            AttributeBlockCollection ReturnTypeAttributes = null;
            TypeParameterCollection TypeParameters = null;

            ValidateModifierList(modifiers, ValidModifiers);

            DelegateLocation = ReadLocation();

            if (Peek().Type == TokenType.Sub)
            {
                SubOrFunctionLocation = ReadLocation();
                MethodType = TreeType.SubDeclaration;
            }
            else if (Peek().Type == TokenType.Function)
            {
                SubOrFunctionLocation = ReadLocation();
                MethodType = TreeType.FunctionDeclaration;
            }
            else
            {
                ReportSyntaxError(SyntaxErrorType.ExpectedSubOrFunction, Peek());
                MethodType = TreeType.SubDeclaration;
            }

            Name = ParseSimpleName(false);

            if (ErrorInConstruct)
            {
                ResyncAt(TokenType.LeftParenthesis, TokenType.As);
            }

            TypeParameters = ParseTypeParameters();

            if (ErrorInConstruct)
            {
                ResyncAt(TokenType.LeftParenthesis, TokenType.As);
            }

            Parameters = ParseParameters();

            if (MethodType == TreeType.FunctionDeclaration && Peek().Type == TokenType.As)
            {
                AsLocation = ReadLocation();
                ReturnTypeAttributes = ParseAttributes();
                ReturnType = ParseTypeName(true);

                if (ErrorInConstruct)
                {
                    ResyncAt();
                }
            }

            if (MethodType == TreeType.SubDeclaration)
            {
                return new DelegateSubDeclaration(attributes, modifiers, DelegateLocation, SubOrFunctionLocation, Name, TypeParameters, Parameters, SpanFrom(startLocation), ParseTrailingComments());
            }
            else
            {
                return new DelegateFunctionDeclaration(attributes, modifiers, DelegateLocation, SubOrFunctionLocation, Name, TypeParameters, Parameters, AsLocation, ReturnTypeAttributes, ReturnType,
                SpanFrom(startLocation), ParseTrailingComments());
            }
        }

        private Declaration ParseTypeListDeclaration(Location startLocation, AttributeBlockCollection attributes, ModifierCollection modifiers, TreeType listType)
        {
            List<Location> CommaLocations = new List<Location>();
            List<TypeName> Types = new List<TypeName>();
            Token ListStart;

            Read();

            if (attributes != null)
                ReportSyntaxError(SyntaxErrorType.SpecifiersInvalidOnTypeListDeclaration, attributes.Span);

            if (modifiers != null)
                ReportSyntaxError(SyntaxErrorType.SpecifiersInvalidOnTypeListDeclaration, modifiers.Span);

            ListStart = Peek();

            do
            {
                if (Types.Count > 0)
                    CommaLocations.Add(ReadLocation());

                Types.Add(ParseTypeName(false));

                if (ErrorInConstruct)
                    ResyncAt(TokenType.Comma);
            }
            while (Peek().Type == TokenType.Comma);

            if (listType == TreeType.InheritsDeclaration)
                return new InheritsDeclaration(new TypeNameCollection(Types, CommaLocations, SpanFrom(ListStart)), SpanFrom(startLocation), ParseTrailingComments());
            else
                return new ImplementsDeclaration(new TypeNameCollection(Types, CommaLocations, SpanFrom(ListStart)), SpanFrom(startLocation), ParseTrailingComments());
        }

        private Declaration ParseNamespaceDeclaration(Location startLocation, AttributeBlockCollection attributes, ModifierCollection modifiers)
        {
            Location KeywordLocation;
            Name Name;
            DeclarationCollection Members;
            EndBlockDeclaration EndBlockDeclaration = null;
            List<Comment> Comments = null;

            if (attributes != null)
            {
                ReportSyntaxError(SyntaxErrorType.SpecifiersInvalidOnNamespaceDeclaration, attributes.Span);
            }

            if (modifiers != null)
            {
                ReportSyntaxError(SyntaxErrorType.SpecifiersInvalidOnNamespaceDeclaration, modifiers.Span);
            }

            KeywordLocation = ReadLocation();

            Name = ParseName(false);

            if (ErrorInConstruct)
            {
                ResyncAt();
            }

            Members = ParseDeclarationBlock(SpanFrom(startLocation), TreeType.NamespaceDeclaration, ref Comments, ref EndBlockDeclaration);

            return new NamespaceDeclaration(attributes, modifiers, KeywordLocation, Name, Members, EndBlockDeclaration, SpanFrom(startLocation), Comments);
        }

        private Declaration ParseImportsDeclaration(Location startLocation, AttributeBlockCollection attributes, ModifierCollection modifiers)
        {
            List<Import> ImportMembers = new List<Import>();
            List<Location> CommaLocations = new List<Location>();
            Token ListStart;

            if (attributes != null)
            {
                ReportSyntaxError(SyntaxErrorType.SpecifiersInvalidOnImportsDeclaration, attributes.Span);
            }

            if (modifiers != null)
            {
                ReportSyntaxError(SyntaxErrorType.SpecifiersInvalidOnImportsDeclaration, modifiers.Span);
            }

            Read();
            ListStart = Peek();

            do
            {
                if (ImportMembers.Count > 0)
                {
                    CommaLocations.Add(ReadLocation());
                }

                if (PeekAheadFor(TokenType.Equals, TokenType.Comma, TokenType.Period) == TokenType.Equals)
                {
                    Token ImportStart = Peek();
                    SimpleName Name;
                    Location EqualsLocation;
                    TypeName AliasedTypeName;

                    Name = ParseSimpleName(false);
                    EqualsLocation = ReadLocation();
                    AliasedTypeName = ParseNamedTypeName(false);

                    if (ErrorInConstruct)
                    {
                        ResyncAt();
                    }

                    ImportMembers.Add(new AliasImport(Name, EqualsLocation, AliasedTypeName, SpanFrom(ImportStart)));
                }
                else
                {
                    Token ImportStart = Peek();
                    TypeName TypeName;

                    TypeName = ParseNamedTypeName(false);

                    if (ErrorInConstruct)
                    {
                        ResyncAt();
                    }

                    ImportMembers.Add(new NameImport(TypeName, SpanFrom(ImportStart)));
                }
            }
            while (Peek().Type == TokenType.Comma);

            return new ImportsDeclaration(new ImportCollection(ImportMembers, CommaLocations, SpanFrom(ListStart)), SpanFrom(startLocation), ParseTrailingComments());
        }

        private Declaration ParseOptionDeclaration(Location startLocation, AttributeBlockCollection attributes, ModifierCollection modifiers)
        {
            OptionType OptionType;
            Location OptionTypeLocation = Location.Empty;
            Location OptionArgumentLocation = Location.Empty;

            Read();

            if (attributes != null)
                ReportSyntaxError(SyntaxErrorType.SpecifiersInvalidOnOptionDeclaration, attributes.Span);
            
            if (modifiers != null)
                ReportSyntaxError(SyntaxErrorType.SpecifiersInvalidOnOptionDeclaration, modifiers.Span);
            
            if (Peek().AsUnreservedKeyword() == TokenType.Explicit)
            {
                OptionTypeLocation = ReadLocation();

                if (Peek().AsUnreservedKeyword() == TokenType.Off)
                {
                    OptionArgumentLocation = ReadLocation();
                    OptionType = OptionType.ExplicitOff;
                }
                else if (Peek().Type == TokenType.On)
                {
                    OptionArgumentLocation = ReadLocation();
                    OptionType = OptionType.ExplicitOn;
                }
                else if (Peek().Type == TokenType.Identifier)
                {
                    OptionType = OptionType.SyntaxError;
                    ReportSyntaxError(SyntaxErrorType.InvalidOptionExplicitType, SpanFrom(startLocation));
                }
                else
                {
                    OptionType = OptionType.Explicit;
                }
            }
            else if (Peek().AsUnreservedKeyword() == TokenType.Strict)
            {
                OptionTypeLocation = ReadLocation();

                if (Peek().AsUnreservedKeyword() == TokenType.Off)
                {
                    OptionArgumentLocation = ReadLocation();
                    OptionType = OptionType.StrictOff;
                }
                else if (Peek().Type == TokenType.On)
                {
                    OptionArgumentLocation = ReadLocation();
                    OptionType = OptionType.StrictOn;
                }
                else if (Peek().Type == TokenType.Identifier)
                {
                    OptionType = OptionType.SyntaxError;
                    ReportSyntaxError(SyntaxErrorType.InvalidOptionStrictType, SpanFrom(startLocation));
                }
                else
                {
                    OptionType = OptionType.Strict;
                }
            }
            else if (Peek().AsUnreservedKeyword() == TokenType.Compare)
            {
                OptionTypeLocation = ReadLocation();

                if (Peek().AsUnreservedKeyword() == TokenType.Binary)
                {
                    OptionArgumentLocation = ReadLocation();
                    OptionType = OptionType.CompareBinary;
                }
                else if (Peek().AsUnreservedKeyword() == TokenType.Text)
                {
                    OptionArgumentLocation = ReadLocation();
                    OptionType = OptionType.CompareText;
                }
                else
                {
                    OptionType = OptionType.SyntaxError;
                    ReportSyntaxError(SyntaxErrorType.InvalidOptionCompareType, SpanFrom(startLocation));
                }
            }
            else if (Peek().AsUnreservedKeyword() == TokenType.Base)
            {
                OptionType = OptionType.SyntaxError;
                OptionTypeLocation = ReadLocation();
                Token Index = Peek();
                
                if (Index.Type == TokenType.IntegerLiteral)
                {
                    OptionArgumentLocation = ReadLocation();
                    long IndexLiteral = ((IntegerLiteralToken)Index).Literal;
                    
                    if (IndexLiteral == 0)
                        OptionType = OptionType.BaseZero;
                    else if (IndexLiteral == 1)
                        OptionType = OptionType.BaseOne;
                }

                if (OptionType == OptionType.SyntaxError)
                    ReportSyntaxError(SyntaxErrorType.InvalidOptionBaseType, SpanFrom(startLocation));
            }
            else
            {
                OptionType = OptionType.SyntaxError;
                ReportSyntaxError(SyntaxErrorType.InvalidOptionType, SpanFrom(startLocation));
            }

            if (ErrorInConstruct)
                ResyncAt();
            
            return new OptionDeclaration(OptionType, OptionTypeLocation, OptionArgumentLocation, SpanFrom(startLocation), ParseTrailingComments());
        }

        private Declaration ParseAttributeDeclaration()
        {
            AttributeBlockCollection Attributes;

            Attributes = ParseAttributes(AttributeTypes.Module | AttributeTypes.Assembly);

            return new AttributeDeclaration(Attributes, Attributes.Span, ParseTrailingComments());
        }

        private Declaration ParseDeclaration()
        {
            Token empty = null;
            return ParseDeclaration(ref empty);
        }

        private Declaration ParseDeclaration(ref Token terminator)
        {
            Token Start;
            Location StartLocation;
            Declaration Declaration = null;
            AttributeBlockCollection Attributes = null;
            ModifierCollection Modifiers = null;
            TokenType LookAhead = TokenType.None;

            if (AtBeginningOfLine)
            {
                while (ParsePreprocessorStatement(false))
                {
                    // Loop
                }
            }

            Start = Peek();

            ErrorInConstruct = false;

            StartLocation = Peek().Span.Start;
            LookAhead = PeekAheadFor(TokenType.Assembly, TokenType.Module, TokenType.GreaterThan);
            if (Peek().Type != TokenType.LessThan || (LookAhead != TokenType.Assembly && LookAhead != TokenType.Module))
            {
                Attributes = ParseAttributes();
                Modifiers = ParseDeclarationModifierList();
            }

            Token Current = Peek();
            TokenType DeclarationType = Current.Type;
            
            switch (DeclarationType)
            {
                case TokenType.End:
                    if (Attributes == null && Modifiers == null)
                        Declaration = ParseEndDeclaration();
                    else
                        goto Identifier;
            
                    break;

                case TokenType.Property:
                    Declaration = ParsePropertyDeclaration(StartLocation, Attributes, Modifiers);
                    break;

                case TokenType.Declare:
                    Declaration = ParseExternalDeclaration(StartLocation, Attributes, Modifiers);
                    break;

                case TokenType.Sub:
                case TokenType.Function:
                    Declaration = ParseMethodDeclaration(StartLocation, Attributes, Modifiers);
                    break;

                case TokenType.Get:
                case TokenType.Set:
                    if (ImplementsVB60)
                        ReportSyntaxError(SyntaxErrorType.ExpectedIdentifier, Peek());
                    else
                        Declaration = ParseAccessorDeclaration(StartLocation, Attributes, Modifiers);
                    break;

                case TokenType.AddHandler:
                case TokenType.RemoveHandler:
                case TokenType.RaiseEvent:
                    Declaration = ParseEventAccessorDeclaration(StartLocation, Attributes, Modifiers);
                    break;

                case TokenType.Event:
                    Declaration = ParseEventDeclaration(StartLocation, Attributes, Modifiers);
                    break;

                case TokenType.Operator:
                    Declaration = ParseOperatorDeclaration(StartLocation, Attributes, Modifiers);
                    break;

                case TokenType.Delegate:
                    Declaration = ParseDelegateDeclaration(StartLocation, Attributes, Modifiers);
                    break;

                case TokenType.Class:
                    Declaration = ParseTypeDeclaration(StartLocation, Attributes, Modifiers, TreeType.ClassDeclaration);
                    break;

                case TokenType.Structure:
                case TokenType.Type:
                    Declaration = ParseTypeDeclaration(StartLocation, Attributes, Modifiers, TreeType.StructureDeclaration);
                    break;
                
                case TokenType.Module:
                    Declaration = ParseTypeDeclaration(StartLocation, Attributes, Modifiers, TreeType.ModuleDeclaration);
                    break;

                case TokenType.Interface:
                    Declaration = ParseTypeDeclaration(StartLocation, Attributes, Modifiers, TreeType.InterfaceDeclaration);
                    break;

                case TokenType.Enum:
                    Declaration = ParseEnumDeclaration(StartLocation, Attributes, Modifiers);
                    break;

                case TokenType.Namespace:
                    Declaration = ParseNamespaceDeclaration(StartLocation, Attributes, Modifiers);
                    break;

                case TokenType.Implements:
                    Declaration = ParseTypeListDeclaration(StartLocation, Attributes, Modifiers, TreeType.ImplementsDeclaration);
                    break;

                case TokenType.Inherits:
                    Declaration = ParseTypeListDeclaration(StartLocation, Attributes, Modifiers, TreeType.InheritsDeclaration);
                    break;

                case TokenType.Imports:
                    Declaration = ParseImportsDeclaration(StartLocation, Attributes, Modifiers);
                    break;

                case TokenType.Option:
                    Declaration = ParseOptionDeclaration(StartLocation, Attributes, Modifiers);
                    break;

                case TokenType.LessThan:
                    Declaration = ParseAttributeDeclaration();
                    break;

                case TokenType.Identifier:
                Identifier:
                    if (Peek().AsUnreservedKeyword() == TokenType.Custom && PeekAheadOne().Type == TokenType.Event)
                        Declaration = ParseCustomEventDeclaration(StartLocation, Attributes, Modifiers);
                    else
                        Declaration = ParseVariableListDeclaration(StartLocation, Attributes, Modifiers);

                    break;

                case TokenType.LineTerminator:
                case TokenType.Colon:
                case TokenType.EndOfStream:
                    if (Attributes != null || Modifiers != null)
                        ReportSyntaxError(SyntaxErrorType.ExpectedIdentifier, Peek());
                    
                    break;

                case TokenType.Comment:
                    List<Comment> Comments = new List<Comment>();
                    Token LastTerminator;

                    do
                    {
                        CommentToken CommentToken = (CommentToken)Scanner.Read();
                        Comments.Add(new Comment(CommentToken.Comment, CommentToken.IsREM, CommentToken.Span));
                        LastTerminator = Read();
                        // Eat the terminator of the comment
                    }
                    while (Peek().Type == TokenType.Comment);

                    GoBack(LastTerminator);

                    Declaration = new EmptyDeclaration(SpanFrom(Start), Comments);
                    break;

                default:
                    ReportSyntaxError(SyntaxErrorType.SyntaxError, Peek());
                    break;
            }

            terminator = VerifyEndOfStatement();

            return Declaration;
        }

        private Declaration ParseDeclarationInEnum()
        {
            Token empty = null;
            return ParseDeclarationInEnum(ref empty);
        }

        private Declaration ParseDeclarationInEnum(ref Token terminator)
        {
            Token Start = null;
            Location StartLocation = Location.Empty;
            AttributeBlockCollection Attributes = null;
            SimpleName Name = null;
            Location EqualsLocation = Location.Empty;
            Expression Expression = null;
            Declaration Declaration = null;

            if (AtBeginningOfLine)
            {
                while (ParsePreprocessorStatement(false))
                {
                    // Loop
                }
            }

            Start = Peek();

            if (Start.Type == TokenType.Comment)
            {
                List<Comment> Comments = new List<Comment>();
                Token LastTerminator;

                do
                {
                    CommentToken CommentToken = (CommentToken)Scanner.Read();
                    Comments.Add(new Comment(CommentToken.Comment, CommentToken.IsREM, CommentToken.Span));
                    LastTerminator = Read();
                    // Eat the terminator of the comment
                }
                while (Peek().Type == TokenType.Comment);
                GoBack(LastTerminator);

                Declaration = new EmptyDeclaration(SpanFrom(Start), Comments);
                goto HaveStatement;
            }

            if (Start.Type == TokenType.LineTerminator || Start.Type == TokenType.Colon)
            {
                goto HaveStatement;
            }

            ErrorInConstruct = false;

            StartLocation = Peek().Span.Start;
            Attributes = ParseAttributes();

            if (Peek().Type == TokenType.End && Attributes == null)
            {
                Declaration = ParseEndDeclaration();
                goto HaveStatement;
            }

            Name = ParseSimpleName(false);

            if (ErrorInConstruct)
            {
                ResyncAt(TokenType.Equals);
            }

            if (Peek().Type == TokenType.Equals)
            {
                EqualsLocation = ReadLocation();
                Expression = ParseExpression();

                if (ErrorInConstruct)
                {
                    ResyncAt();
                }
            }

            Declaration = new EnumValueDeclaration(Attributes, Name, EqualsLocation, Expression, SpanFrom(StartLocation), ParseTrailingComments());
        HaveStatement:

            terminator = VerifyEndOfStatement();

            return Declaration;
        }

        private DeclarationCollection ParseDeclarationBlock(Span blockStartSpan, TreeType blockType, ref List<Comment> Comments)
        {

            EndBlockDeclaration empty = null;
            return ParseDeclarationBlock(blockStartSpan, blockType, ref Comments, ref empty);
        }

        private DeclarationCollection ParseDeclarationBlock(Span blockStartSpan, TreeType blockType, ref List<Comment> Comments, ref EndBlockDeclaration endDeclaration)
        {
            List<Declaration> Declarations = new List<Declaration>();
            List<Location> ColonLocations = new List<Location>();
            Token Terminator;
            Token Start;
            Location DeclarationsEnd;
            bool BlockTerminated = false;

            Comments = ParseTrailingComments();
            Terminator = VerifyEndOfStatement();

            if (Terminator.Type == TokenType.Colon)
                ColonLocations.Add(Terminator.Span.Start);
            
            Start = Peek();
            DeclarationsEnd = Start.Span.Finish;
            endDeclaration = null;

            PushBlockContext(blockType);

            while (Peek().Type != TokenType.EndOfStream)
            {
                Token PreviousTerminator = Terminator;
                Declaration Declaration;

                if (blockType == TreeType.EnumDeclaration)
                    Declaration = ParseDeclarationInEnum(ref Terminator);
                else
                    Declaration = ParseDeclaration(ref Terminator);
                
                if (Declaration != null)
                {
                    SyntaxErrorType ErrorType = SyntaxErrorType.None;

                    if (Declaration.Type == TreeType.EndBlockDeclaration)
                    {
                        EndBlockDeclaration PotentialEndDeclaration = (EndBlockDeclaration)Declaration;

                        if (DeclarationEndsBlock(blockType, PotentialEndDeclaration))
                        {
                            endDeclaration = PotentialEndDeclaration;
                            GoBack(Terminator);
                            BlockTerminated = true;
                            break;
                        }
                        else
                        {
                            bool DeclarationEndsOuterBlock = false;

                            // If the end Declaration matches an outer block context, then we want to unwind
                            // up to that level. Otherwise, we want to just give an error and keep going.
                            foreach (TreeType BlockContext in BlockContextStack)
                            {
                                if (DeclarationEndsBlock(BlockContext, PotentialEndDeclaration))
                                {
                                    DeclarationEndsOuterBlock = true;
                                    break;
                                }
                            }

                            if (DeclarationEndsOuterBlock)
                            {
                                ReportMismatchedEndError(blockType, Declaration.Span);
                                // CONSIDER: Can we avoid parsing and re-parsing this declaration?
                                GoBack(PreviousTerminator);
                                // We consider the block terminated.
                                BlockTerminated = true;
                                break;
                            }
                            else
                            {
                                ReportMissingBeginDeclarationError(PotentialEndDeclaration);
                            }
                        }
                    }
                    else
                    {
                        ErrorType = ValidDeclaration(blockType, Declaration, Declarations);

                        if (ErrorType != SyntaxErrorType.None)
                            ReportSyntaxError(ErrorType, Declaration.Span);
                    }

                    Declarations.Add(Declaration);
                }

                if (Terminator.Type == TokenType.Colon)
                    ColonLocations.Add(Terminator.Span.Start);
                
                DeclarationsEnd = Terminator.Span.Finish;
            }

            if (!BlockTerminated)
                ReportMismatchedEndError(blockType, blockStartSpan);
           
            PopBlockContext();

            NormalizeDeclarationList(Declarations);

            if (Declarations.Count == 0 && ColonLocations.Count == 0)
                return null;
            else
                return new DeclarationCollection(Declarations, ColonLocations, new Span(Start.Span.Start, DeclarationsEnd));
        }

        #endregion

        #region Parameters

        private Parameter ParseParameter()
        {
            Token Start = Peek();
            AttributeBlockCollection Attributes = null;
            ModifierCollection Modifiers = null;
            VariableName VariableName = null;
            Location AsLocation = Location.Empty;
            TypeName Type = null;
            Location EqualsLocation = Location.Empty;
            Initializer Initializer = null;

            Attributes = ParseAttributes();
            Modifiers = ParseParameterModifierList();
            VariableName = ParseVariableName(false);

            if (ErrorInConstruct)
            {
                // If we see As before a comma or RParen, then assume that
                // we are still on the same parameter. Otherwise, don't resync
                // and allow the caller to decide how to recover.
                if (PeekAheadFor(TokenType.As, TokenType.Comma, TokenType.RightParenthesis) == TokenType.As)
                    ResyncAt(TokenType.As);
            }

            if (Peek().Type == TokenType.As)
            {
                AsLocation = ReadLocation();
                Type = ParseTypeName(true);
            }

            if (ErrorInConstruct)
                ResyncAt(TokenType.Equals, TokenType.Comma, TokenType.RightParenthesis);

            if (Peek().Type == TokenType.Equals)
            {
                EqualsLocation = ReadLocation();
                Initializer = ParseInitializer();
            }

            if (ErrorInConstruct)
                ResyncAt(TokenType.Comma, TokenType.RightParenthesis);

            return new Parameter(Attributes, Modifiers, VariableName, AsLocation, Type, EqualsLocation, Initializer, SpanFrom(Start));
        }

        private bool ParametersContinue()
        {
            Token NextToken = Peek();

            if (NextToken.Type == TokenType.Comma)
            {
                return true;
            }
            else if (NextToken.Type == TokenType.RightParenthesis || IsEndOfStatement(NextToken))
            {
                return false;
            }

            ReportSyntaxError(SyntaxErrorType.ParameterSyntax, NextToken);
            ResyncAt(TokenType.Comma, TokenType.RightParenthesis);

            if (Peek().Type == TokenType.Comma)
            {
                ErrorInConstruct = false;
                return true;
            }

            return false;
        }

        private ParameterCollection ParseParameters()
        {
            Token Start = Peek();
            List<Parameter> Parameters = new List<Parameter>();
            List<Location> CommaLocations = new List<Location>();
            Location RightParenthesisLocation = Location.Empty;

            if (Start.Type != TokenType.LeftParenthesis)
            {
                return null;
            }
            else
            {
                Read();
            }

            if (Peek().Type != TokenType.RightParenthesis)
            {
                do
                {
                    if (Parameters.Count > 0)
                    {
                        CommaLocations.Add(ReadLocation());
                    }

                    Parameters.Add(ParseParameter());

                    if (ErrorInConstruct)
                    {
                        ResyncAt(TokenType.Comma, TokenType.RightParenthesis);
                    }
                }
                while (ParametersContinue());
            }

            if (Peek().Type == TokenType.RightParenthesis)
            {
                RightParenthesisLocation = ReadLocation();
            }
            else
            {
                Token CurrentToken = Peek();

                // On error, peek for ")" with "(". If ")" seen before 
                // "(", then sync on that. Otherwise, assume missing ")"
                // and let caller decide.
                ResyncAt(TokenType.LeftParenthesis, TokenType.RightParenthesis);

                if (Peek().Type == TokenType.RightParenthesis)
                {
                    ReportSyntaxError(SyntaxErrorType.SyntaxError, Peek());
                    RightParenthesisLocation = ReadLocation();
                }
                else
                {
                    GoBack(CurrentToken);
                    ReportSyntaxError(SyntaxErrorType.ExpectedRightParenthesis, Peek());
                }
            }

            return new ParameterCollection(Parameters, CommaLocations, RightParenthesisLocation, SpanFrom(Start));
        }

        #endregion

        #region Type Parameters

        private TypeConstraintCollection ParseTypeConstraints()
        {
            Token Start = Peek();
            List<Location> CommaLocations = new List<Location>();
            List<TypeName> Types = new List<TypeName>();
            Location RightBracketLocation = Location.Empty;

            if (Peek().Type == TokenType.LeftCurlyBrace)
            {
                Read();

                do
                {
                    if (Types.Count > 0)
                    {
                        CommaLocations.Add(ReadLocation());
                    }

                    Types.Add(ParseTypeName(true));

                    if (ErrorInConstruct)
                    {
                        ResyncAt(TokenType.Comma);
                    }
                }
                while (Peek().Type == TokenType.Comma);

                RightBracketLocation = VerifyExpectedToken(TokenType.RightCurlyBrace);
            }
            else
            {
                Types.Add(ParseTypeName(true));
            }

            return new TypeConstraintCollection(Types, CommaLocations, RightBracketLocation, SpanFrom(Start));
        }

        private TypeParameter ParseTypeParameter()
        {
            Token Start = Peek();
            SimpleName Name = null;
            Location AsLocation = Location.Empty;
            TypeConstraintCollection TypeConstraints = null;

            Name = ParseSimpleName(false);

            if (ErrorInConstruct)
            {
                // If we see As before a comma or RParen, then assume that
                // we are still on the same parameter. Otherwise, don't resync
                // and allow the caller to decide how to recover.
                if (PeekAheadFor(TokenType.As, TokenType.Comma, TokenType.RightParenthesis) == TokenType.As)
                {
                    ResyncAt(TokenType.As);
                }
            }

            if (Peek().Type == TokenType.As)
            {
                AsLocation = ReadLocation();
                TypeConstraints = ParseTypeConstraints();
            }

            if (ErrorInConstruct)
            {
                ResyncAt(TokenType.Equals, TokenType.Comma, TokenType.RightParenthesis);
            }

            if (ErrorInConstruct)
            {
                ResyncAt(TokenType.Comma, TokenType.RightParenthesis);
            }

            return new TypeParameter(Name, AsLocation, TypeConstraints, SpanFrom(Start));
        }

        private TypeParameterCollection ParseTypeParameters()
        {
            Token Start = Peek();
            Location OfLocation;
            List<TypeParameter> TypeParameters = new List<TypeParameter>();
            List<Location> CommaLocations = new List<Location>();
            Location RightParenthesisLocation;

            if (Start.Type != TokenType.LeftParenthesis || ImplementsVB71)
            {
                return null;
            }
            else
            {
                Read();

                if (Peek().Type != TokenType.Of || ImplementsVB71)
                {
                    GoBack(Start);
                    return null;
                }
            }

            OfLocation = VerifyExpectedToken(TokenType.Of);

            do
            {
                if (TypeParameters.Count > 0)
                    CommaLocations.Add(ReadLocation());

                TypeParameters.Add(ParseTypeParameter());

                if (ErrorInConstruct)
                    ResyncAt(TokenType.Comma, TokenType.RightParenthesis);
            }
            while (ParametersContinue());

            RightParenthesisLocation = VerifyExpectedToken(TokenType.RightParenthesis);

            return new TypeParameterCollection(OfLocation, TypeParameters, CommaLocations, RightParenthesisLocation, SpanFrom(Start));
        }

        #endregion

        #region Files

        private FileTree ParseFile()
        {
            List<Declaration> Declarations = new List<Declaration>();
            List<Location> ColonLocations = new List<Location>();
            Token Terminator = null;
            Token Start = Peek();

            while (Peek().Type != TokenType.EndOfStream)
            {
                Declaration Declaration;

                Declaration = ParseDeclaration(ref Terminator);

                if (Declaration != null)
                {
                    SyntaxErrorType ErrorType = SyntaxErrorType.None;

                    ErrorType = ValidDeclaration(TreeType.FileTree, Declaration, Declarations);

                    if (ErrorType != SyntaxErrorType.None)
                        ReportSyntaxError(ErrorType, Declaration.Span);

                    Declarations.Add(Declaration);
                }

                if (Terminator.Type == TokenType.Colon)
                    ColonLocations.Add(Terminator.Span.Start);
            }

            NormalizeDeclarationList(Declarations);

            if (Declarations.Count == 0 && ColonLocations.Count == 0)
                return new FileTree(null, SpanFrom(Start));
            else
                return new FileTree(new DeclarationCollection(Declarations, ColonLocations, SpanFrom(Start)), SpanFrom(Start));
        }

        #endregion

        #region Preprocessor statements

        private void ParseExternalSourceStatement(Token start)
        {
            long Line;
            string File;

            // Consume the ExternalSource keyword
            Read();

            if (CurrentExternalSourceContext != null)
            {
                ResyncAt();
                ReportSyntaxError(SyntaxErrorType.NestedExternalSourceStatement, SpanFrom(start));
            }
            else
            {
                VerifyExpectedToken(TokenType.LeftParenthesis);

                if (Peek().Type != TokenType.StringLiteral)
                {
                    ReportSyntaxError(SyntaxErrorType.ExpectedStringLiteral, Peek());
                    ResyncAt();
                    return;
                }

                File = ((StringLiteralToken)Read()).Literal;
                VerifyExpectedToken(TokenType.Comma);

                if (Peek().Type != TokenType.IntegerLiteral)
                {
                    ReportSyntaxError(SyntaxErrorType.ExpectedIntegerLiteral, Peek());
                    ResyncAt();
                    return;
                }

                Line = ((IntegerLiteralToken)Read()).Literal;

                VerifyExpectedToken(TokenType.RightParenthesis);

                CurrentExternalSourceContext = new ExternalSourceContext();
                {
                    CurrentExternalSourceContext.File = File;
                    CurrentExternalSourceContext.Line = Line;
                    CurrentExternalSourceContext.Start = Peek().Span.Start;
                }
            }
        }

        private void ParseExternalChecksumStatement()
        {
            string Filename;
            string Guid;
            string Checksum;

            // Consume the ExternalChecksum keyword
            Read();
            VerifyExpectedToken(TokenType.LeftParenthesis);

            if (Peek().Type != TokenType.StringLiteral)
            {
                ReportSyntaxError(SyntaxErrorType.ExpectedStringLiteral, Peek());
                ResyncAt();
                return;
            }

            Filename = ((StringLiteralToken)Read()).Literal;
            VerifyExpectedToken(TokenType.Comma);

            if (Peek().Type != TokenType.StringLiteral)
            {
                ReportSyntaxError(SyntaxErrorType.ExpectedStringLiteral, Peek());
                ResyncAt();
                return;
            }

            Guid = ((StringLiteralToken)Read()).Literal;
            VerifyExpectedToken(TokenType.Comma);

            if (Peek().Type != TokenType.StringLiteral)
            {
                ReportSyntaxError(SyntaxErrorType.ExpectedStringLiteral, Peek());
                ResyncAt();
                return;
            }

            Checksum = ((StringLiteralToken)Read()).Literal;
            VerifyExpectedToken(TokenType.RightParenthesis);

            if (ExternalChecksums != null)
            {
                ExternalChecksums.Add(new ExternalChecksum(Filename, Guid, Checksum));
            }
        }

        private void ParseRegionStatement(Token start, bool statementLevel)
        {
            string Description;
            RegionContext RegionContext;

            if (statementLevel == true)
            {
                ResyncAt();
                ReportSyntaxError(SyntaxErrorType.RegionInsideMethod, SpanFrom(start));
                return;
            }

            // Consume the Region keyword
            Read();

            if (Peek().Type != TokenType.StringLiteral)
            {
                ReportSyntaxError(SyntaxErrorType.ExpectedStringLiteral, Peek());
                ResyncAt();
                return;
            }

            Description = ((StringLiteralToken)Read()).Literal;

            RegionContext = new RegionContext();
            RegionContext.Description = Description;
            RegionContext.Start = Peek().Span.Start;
            RegionContextStack.Push(RegionContext);
        }

        private void ParseEndPreprocessingStatement(Token start, bool statementLevel)
        {
            // Consume the End keyword
            Read();

            if (Peek().AsUnreservedKeyword() == TokenType.ExternalSource)
            {
                Read();

                if (CurrentExternalSourceContext == null)
                {
                    ReportSyntaxError(SyntaxErrorType.EndExternalSourceWithoutExternalSource, SpanFrom(start));
                    ResyncAt();
                }
                else
                {
                    if (ExternalLineMappings != null)
                    {
                        {
                            ExternalLineMappings.Add(new ExternalLineMapping(CurrentExternalSourceContext.Start, start.Span.Start, CurrentExternalSourceContext.File, CurrentExternalSourceContext.Line));
                        }
                    }
                    CurrentExternalSourceContext = null;
                }

                return;
            }
            else if (Peek().AsUnreservedKeyword() == TokenType.Region)
            {
                Read();

                if (statementLevel == true)
                {
                    ResyncAt();
                    ReportSyntaxError(SyntaxErrorType.RegionInsideMethod, SpanFrom(start));
                    return;
                }

                if (RegionContextStack.Count == 0)
                {
                    ReportSyntaxError(SyntaxErrorType.EndRegionWithoutRegion, SpanFrom(start));
                    ResyncAt();
                }
                else
                {
                    RegionContext RegionContext = RegionContextStack.Pop();

                    if (SourceRegions != null)
                    {
                        SourceRegions.Add(new SourceRegion(RegionContext.Start, start.Span.Start, RegionContext.Description));
                    }
                }

                return;
            }
            else if (Peek().Type == TokenType.If)
            {
                // Read the If keyword
                Read();

                if (ConditionalCompilationContextStack.Count == 0)
                {
                    ReportSyntaxError(SyntaxErrorType.CCEndIfWithoutCCIf, SpanFrom(start));
                }
                else
                {
                    ConditionalCompilationContextStack.Pop();
                }

                return;
            }

            ResyncAt();
            ReportSyntaxError(SyntaxErrorType.ExpectedEndKind, Peek());
        }

        private static object EvaluateCCLiteral(Expression expression)
        {
            switch (expression.Type)
            {
                case TreeType.IntegerLiteralExpression:
                    return ((IntegerLiteralExpression)expression).Literal;

                case TreeType.FloatingPointLiteralExpression:
                    return ((FloatingPointLiteralExpression)expression).Literal;

                case TreeType.StringLiteralExpression:
                    return ((StringLiteralExpression)expression).Literal;

                case TreeType.CharacterLiteralExpression:
                    return ((CharacterLiteralExpression)expression).Literal;

                case TreeType.DateLiteralExpression:
                    return ((DateLiteralExpression)expression).Literal;

                case TreeType.DecimalLiteralExpression:
                    return ((DecimalLiteralExpression)expression).Literal;

                case TreeType.BooleanLiteralExpression:
                    return ((BooleanLiteralExpression)expression).Literal;

                default:
                    Debug.Fail("Unexpected!");
                    return null;
            }
        }

        private static TypeCode TypeCodeOfCastExpression(IntrinsicType castType)
        {
            switch (castType)
            {
                case IntrinsicType.Boolean:
                    return TypeCode.Boolean;

                case IntrinsicType.Byte:
                    return TypeCode.Byte;

                case IntrinsicType.Char:
                    return TypeCode.Char;

                case IntrinsicType.Date:
                    return TypeCode.DateTime;

                case IntrinsicType.Decimal:
                    return TypeCode.Decimal;

                case IntrinsicType.Double:
                    return TypeCode.Double;

                case IntrinsicType.Integer:
                    return TypeCode.Int32;

                case IntrinsicType.Long:
                    return TypeCode.Int64;

                case IntrinsicType.Object:
                    return TypeCode.Object;

                case IntrinsicType.Short:
                    return TypeCode.Int16;

                case IntrinsicType.Single:
                    return TypeCode.Single;

                case IntrinsicType.String:
                    return TypeCode.String;

                default:
                    Debug.Fail("Unexpected!");
                    return TypeCode.Empty;
            }
        }

        private object EvaluateCCCast(IntrinsicCastExpression expression)
        {
            // This cast is safe because only intrinsics are ever returned
            IConvertible Operand = (IConvertible)EvaluateCCExpression(expression.Operand);
            TypeCode OperandType;
            TypeCode CastType = TypeCodeOfCastExpression(expression.IntrinsicType);

            if (CastType == TypeCode.Empty)
            {
                return null;
            }

            if (Operand == null)
            {
                Operand = 0;
            }

            OperandType = Operand.GetTypeCode();

            if (CastType == OperandType || CastType == TypeCode.Object)
            {
                return Operand;
            }

            switch (OperandType)
            {
                case TypeCode.Boolean:
                    if (CastType == TypeCode.Byte)
                    {
                        Operand = 255;
                    }
                    else
                    {
                        Operand = -1;
                    }

                    OperandType = TypeCode.Int32;
                    break;

                case TypeCode.String:
                    if (CastType != TypeCode.Char)
                    {
                        ReportSyntaxError(SyntaxErrorType.CantCastStringInCCExpression, expression.Span);
                        return null;
                    }

                    break;

                case TypeCode.Char:
                    if (CastType != TypeCode.String)
                    {
                        ReportSyntaxError(SyntaxErrorType.InvalidCCCast, expression.Span);
                        return null;
                    }

                    break;

                case TypeCode.DateTime:
                    ReportSyntaxError(SyntaxErrorType.InvalidCCCast, expression.Span);
                    return null;
            }

            switch (expression.IntrinsicType)
            {
                case IntrinsicType.Boolean:
                    return (bool)Operand;

                case IntrinsicType.Byte:
                    return (byte)Operand;

                case IntrinsicType.Short:
                    return (short)Operand;

                case IntrinsicType.Integer:
                    return (int)Operand;

                case IntrinsicType.Long:
                    return (long)Operand;

                case IntrinsicType.Decimal:
                    return (decimal)Operand;

                case IntrinsicType.Single:
                    return (float)Operand;

                case IntrinsicType.Double:
                    return (double)Operand;

                case IntrinsicType.Char:
                    if (OperandType == TypeCode.String)
                    {
                        return char.Parse(((string)Operand));
                    }

                    ReportSyntaxError(SyntaxErrorType.InvalidCCCast, expression.Span);
                    return null;

                case IntrinsicType.String:
                    if (OperandType == TypeCode.Char)
                    {
                        return char.Parse((string)Operand).ToString();
                    }

                    ReportSyntaxError(SyntaxErrorType.CantCastStringInCCExpression, expression.Span);
                    return null;

                case IntrinsicType.Date:
                    ReportSyntaxError(SyntaxErrorType.InvalidCCCast, expression.Span);
                    return null;

                default:
                    Debug.Fail("Unexpected!");
                    return null;
            }
        }

        private object EvaluateCCUnaryOperator(UnaryOperatorExpression expression)
        {
            // This cast is safe because only intrinsics are ever returned
            IConvertible Operand = (IConvertible)EvaluateCCExpression(expression.Operand);
            TypeCode OperandType;

            if (Operand == null)
            {
                Operand = 0;
            }

            OperandType = Operand.GetTypeCode();

            if (OperandType == TypeCode.String || OperandType == TypeCode.Char || OperandType == TypeCode.DateTime)
            {
                ReportSyntaxError(SyntaxErrorType.InvalidCCOperator, expression.Span);
                return null;
            }

            switch (expression.Operator)
            {
                case UnaryOperatorType.UnaryPlus:
                    if (OperandType == TypeCode.Boolean)
                    {
                        ReportSyntaxError(SyntaxErrorType.InvalidCCOperator, expression.Span);
                        return null;
                    }
                    else
                    {
                        return Operand;
                    }

                case UnaryOperatorType.Negate:
                    if (OperandType == TypeCode.Boolean || OperandType == TypeCode.Byte)
                    {
                        ReportSyntaxError(SyntaxErrorType.InvalidCCOperator, expression.Span);
                        return null;
                    }
                    else
                    {
                        return ObjectType.NegObj(Operand);
                    }

                case UnaryOperatorType.Not:
                    if (OperandType == TypeCode.Decimal || OperandType == TypeCode.Single || OperandType == TypeCode.Double)
                    {
                        ReportSyntaxError(SyntaxErrorType.InvalidCCOperator, expression.Span);
                        return null;
                    }
                    else
                    {
                        return ObjectType.NotObj(Operand);
                    }

                default:
                    Debug.Fail("Unexpected!");
                    return null;
            }
        }

        private static bool EitherIsTypeCode(TypeCode x, TypeCode y, TypeCode type)
        {
            return x == type || y == type;
        }

        private static bool IsEitherTypeCode(TypeCode x, TypeCode type1, TypeCode type2)
        {
            return x == type1 || x == type2;
        }

        private object EvaluateCCBinaryOperator(BinaryOperatorExpression expression)
        {
            // This cast is safe because only intrinsics are ever returned
            IConvertible LeftOperand = (IConvertible)EvaluateCCExpression(expression.LeftOperand);
            IConvertible RightOperand = (IConvertible)EvaluateCCExpression(expression.RightOperand);
            TypeCode LeftOperandType;
            TypeCode RightOperandType;

            if (LeftOperand == null)
            {
                LeftOperand = 0;
            }

            if (RightOperand == null)
            {
                RightOperand = 0;
            }

            LeftOperandType = LeftOperand.GetTypeCode();
            RightOperandType = RightOperand.GetTypeCode();

            if (EitherIsTypeCode(LeftOperandType, RightOperandType, TypeCode.DateTime))
            {
                ReportSyntaxError(SyntaxErrorType.InvalidCCOperator, expression.Span);
                return null;
            }

            if (expression.Operator != BinaryOperatorType.Concatenate && expression.Operator != BinaryOperatorType.Plus && expression.Operator != BinaryOperatorType.Equals && expression.Operator != BinaryOperatorType.NotEquals && (EitherIsTypeCode(LeftOperandType, RightOperandType, TypeCode.Char) || EitherIsTypeCode(LeftOperandType, RightOperandType, TypeCode.String)))
            {
                ReportSyntaxError(SyntaxErrorType.InvalidCCOperator, expression.Span);
                return null;
            }

            switch (expression.Operator)
            {
                case BinaryOperatorType.Plus:
                    if (EitherIsTypeCode(LeftOperandType, RightOperandType, TypeCode.String) || EitherIsTypeCode(LeftOperandType, RightOperandType, TypeCode.Char))
                    {
                        if (!IsEitherTypeCode(LeftOperandType, TypeCode.Char, TypeCode.String) || !IsEitherTypeCode(LeftOperandType, TypeCode.Char, TypeCode.String))
                        {
                            ReportSyntaxError(SyntaxErrorType.InvalidCCOperator, expression.Span);
                            return null;
                        }
                        else
                        {
                            return (string)LeftOperand + (string)RightOperand;
                        }
                    }
                    else
                    {
                        return ObjectType.AddObj(LeftOperand, RightOperand);
                    }

                case BinaryOperatorType.Minus:
                    return ObjectType.SubObj(LeftOperand, RightOperand);

                case BinaryOperatorType.Multiply:
                    return ObjectType.MulObj(LeftOperand, RightOperand);

                case BinaryOperatorType.IntegralDivide:
                    return ObjectType.IDivObj(LeftOperand, RightOperand);

                case BinaryOperatorType.Divide:
                    return ObjectType.DivObj(LeftOperand, RightOperand);

                case BinaryOperatorType.Modulus:
                    return ObjectType.ModObj(LeftOperand, RightOperand);

                case BinaryOperatorType.Power:
                    return ObjectType.PowObj(LeftOperand, RightOperand);

                case BinaryOperatorType.ShiftLeft:
                    return ObjectType.ShiftLeftObj(LeftOperand, (int)RightOperand);

                case BinaryOperatorType.ShiftRight:
                    return ObjectType.ShiftRightObj(LeftOperand, (int)RightOperand);

                case BinaryOperatorType.And:
                    return ObjectType.BitAndObj(LeftOperand, (int)RightOperand);

                case BinaryOperatorType.Or:
                    return ObjectType.BitOrObj(LeftOperand, (int)RightOperand);

                case BinaryOperatorType.Xor:
                    return ObjectType.BitXorObj(LeftOperand, (int)RightOperand);

                case BinaryOperatorType.AndAlso:
                    return (bool)LeftOperand && (bool)RightOperand;

                case BinaryOperatorType.OrElse:
                    return (bool)LeftOperand || (bool)RightOperand;

                case BinaryOperatorType.Equals:
                    if ((EitherIsTypeCode(LeftOperandType, RightOperandType, TypeCode.String) || EitherIsTypeCode(LeftOperandType, RightOperandType, TypeCode.Char)) && (!IsEitherTypeCode(LeftOperandType, TypeCode.Char, TypeCode.String) || !IsEitherTypeCode(LeftOperandType, TypeCode.Char, TypeCode.String)))
                    {
                        ReportSyntaxError(SyntaxErrorType.InvalidCCOperator, expression.Span);
                        return null;
                    }


                    return ObjectType.ObjTst(LeftOperand, RightOperand, false) == 0;

                case BinaryOperatorType.NotEquals:
                    if ((EitherIsTypeCode(LeftOperandType, RightOperandType, TypeCode.String) || EitherIsTypeCode(LeftOperandType, RightOperandType, TypeCode.Char)) && (!IsEitherTypeCode(LeftOperandType, TypeCode.Char, TypeCode.String) || !IsEitherTypeCode(LeftOperandType, TypeCode.Char, TypeCode.String)))
                    {
                        ReportSyntaxError(SyntaxErrorType.InvalidCCOperator, expression.Span);
                        return null;
                    }


                    return ObjectType.ObjTst(LeftOperand, RightOperand, false) != 0;

                case BinaryOperatorType.LessThan:
                    return ObjectType.ObjTst(LeftOperand, RightOperand, false) == -1;

                case BinaryOperatorType.GreaterThan:
                    return ObjectType.ObjTst(LeftOperand, RightOperand, false) == 1;

                case BinaryOperatorType.LessThanEquals:
                    return ObjectType.ObjTst(LeftOperand, RightOperand, false) != 1;

                case BinaryOperatorType.GreaterThanEquals:
                    return ObjectType.ObjTst(LeftOperand, RightOperand, false) != -1;

                case BinaryOperatorType.Concatenate:
                    if (!IsEitherTypeCode(LeftOperandType, TypeCode.Char, TypeCode.String) || !IsEitherTypeCode(LeftOperandType, TypeCode.Char, TypeCode.String))
                    {
                        ReportSyntaxError(SyntaxErrorType.InvalidCCOperator, expression.Span);
                        return null;
                    }
                    else
                    {
                        return (string)LeftOperand + (string)RightOperand;
                    }

                default:
                    Debug.Fail("Unexpected!");
                    return null;
            }
        }

        private object EvaluateCCExpression(Expression expression)
        {
            switch (expression.Type)
            {
                case TreeType.SyntaxError:
                    break;
                // Do nothing

                case TreeType.NothingExpression:
                    return null;

                case TreeType.IntegerLiteralExpression:
                case TreeType.FloatingPointLiteralExpression:
                case TreeType.StringLiteralExpression:
                case TreeType.CharacterLiteralExpression:
                case TreeType.DateLiteralExpression:
                case TreeType.DecimalLiteralExpression:
                case TreeType.BooleanLiteralExpression:
                    return EvaluateCCLiteral(expression);

                case TreeType.ParentheticalExpression:
                    return EvaluateCCExpression(((ParentheticalExpression)expression).Operand);

                case TreeType.SimpleNameExpression:
                    if (ConditionalCompilationConstants.ContainsKey(((SimpleNameExpression)expression).Name.Name))
                    {
                        return ConditionalCompilationConstants[((SimpleNameExpression)expression).Name.Name];
                    }
                    else
                    {
                        return null;
                    }

                case TreeType.IntrinsicCastExpression:
                    return EvaluateCCCast((IntrinsicCastExpression)expression);

                case TreeType.UnaryOperatorExpression:
                    return EvaluateCCUnaryOperator((UnaryOperatorExpression)expression);

                case TreeType.BinaryOperatorExpression:
                    return EvaluateCCBinaryOperator((BinaryOperatorExpression)expression);

                default:
                    ReportSyntaxError(SyntaxErrorType.CCExpressionRequired, expression.Span);
                    break;
            }

            return null;
        }

        private void ParseConditionalConstantStatement()
        {
            IdentifierToken Identifier;
            Expression Expression;

            // Consume the Const keyword
            Read();

            if (Peek().Type == TokenType.Identifier)
            {
                Identifier = (IdentifierToken)Read();
            }
            else
            {
                ReportSyntaxError(SyntaxErrorType.ExpectedIdentifier, Peek());
                ResyncAt();
                return;
            }

            VerifyExpectedToken(TokenType.Equals);
            Expression = ParseExpression();

            if (!ErrorInConstruct)
            {
                ConditionalCompilationConstants.Add(Identifier.Identifier, EvaluateCCExpression(Expression));
            }
            else
            {
                ResyncAt();
            }
        }

        private void ParseConditionalIfStatement()
        {
            Expression Expression;
            ConditionalCompilationContext CCContext;

            // Consume the If
            Read();

            Expression = ParseExpression();

            if (ErrorInConstruct)
            {
                ResyncAt(TokenType.Then);
            }

            if (Peek().Type == TokenType.Then)
            {
                // Consume the Then keyword
                Read();
            }

            CCContext = new ConditionalCompilationContext();
            {
                CCContext.BlockActive = EvaluateCCExpression(Expression) != null;
                CCContext.AnyBlocksActive = CCContext.BlockActive;
            }
            ConditionalCompilationContextStack.Push(CCContext);
        }

        private void ParseConditionalElseIfStatement(Token start)
        {
            Expression Expression;
            ConditionalCompilationContext CCContext;

            // Consume the If
            Read();

            Expression = ParseExpression();

            if (ErrorInConstruct)
            {
                ResyncAt(TokenType.Then);
            }

            if (Peek().Type == TokenType.Then)
            {
                // Consume the Then keyword
                Read();
            }

            if (ConditionalCompilationContextStack.Count == 0)
            {
                ReportSyntaxError(SyntaxErrorType.CCElseIfWithoutCCIf, SpanFrom(start));
            }
            else
            {
                CCContext = ConditionalCompilationContextStack.Peek();

                if (CCContext.SeenElse)
                {
                    ReportSyntaxError(SyntaxErrorType.CCElseIfAfterCCElse, SpanFrom(start));
                    CCContext.BlockActive = false;
                }
                else if (CCContext.BlockActive)
                {
                    CCContext.BlockActive = false;
                }
                else if (!CCContext.AnyBlocksActive && (bool)EvaluateCCExpression(Expression))
                {
                    CCContext.BlockActive = true;
                    CCContext.AnyBlocksActive = true;
                }
            }
        }

        private void ParseConditionalElseStatement(Token start)
        {
            ConditionalCompilationContext CCContext;

            // Consume the else
            Read();

            if (ConditionalCompilationContextStack.Count == 0)
            {
                ReportSyntaxError(SyntaxErrorType.CCElseWithoutCCIf, SpanFrom(start));
            }
            else
            {
                CCContext = ConditionalCompilationContextStack.Peek();

                if (CCContext.SeenElse)
                {
                    ReportSyntaxError(SyntaxErrorType.CCElseAfterCCElse, SpanFrom(start));
                    CCContext.BlockActive = false;
                }
                else
                {
                    CCContext.SeenElse = true;

                    if (CCContext.BlockActive)
                    {
                        CCContext.BlockActive = false;
                    }
                    else if (!CCContext.AnyBlocksActive)
                    {
                        CCContext.BlockActive = true;
                    }
                }
            }
        }

        private bool ParsePreprocessorStatement(bool statementLevel)
        {
            Token Start = Peek();

            Debug.Assert(AtBeginningOfLine, "Must be at beginning of line!");

            if (!Preprocess)
                return false;

            if (Start.Type == TokenType.Pound)
            {
                ErrorInConstruct = false;

                // Consume the pound
                Read();

                switch (Peek().AsUnreservedKeyword())
                {
                    case TokenType.Const:
                        ParseConditionalConstantStatement();
                        break;

                    case TokenType.If:
                        ParseConditionalIfStatement();
                        break;

                    case TokenType.Else:
                        ParseConditionalElseStatement(Start);
                        break;

                    case TokenType.ElseIf:
                        ParseConditionalElseIfStatement(Start);
                        break;

                    case TokenType.ExternalSource:
                        ParseExternalSourceStatement(Start);
                        break;

                    case TokenType.Region:
                        ParseRegionStatement(Start, statementLevel);
                        break;

                    case TokenType.ExternalChecksum:
                        ParseExternalChecksumStatement();
                        break;

                    case TokenType.End:
                        ParseEndPreprocessingStatement(Start, statementLevel);
                        break;

                    default:
                        ResyncAt();
                        ReportSyntaxError(SyntaxErrorType.InvalidPreprocessorStatement, SpanFrom(Start));
                        break;
                }

                ParseTrailingComments();

                if (Peek().Type != TokenType.LineTerminator && Peek().Type != TokenType.EndOfStream)
                {
                    ReportSyntaxError(SyntaxErrorType.ExpectedEndOfStatement, Peek());
                    ResyncAt();
                }

                Read();
                return true;
            }
            else
            {
                // If we're in a false conditional compilation statement, then keep reading lines as if they
                // were preprocessing statements until we are done.
                if (Start.Type != TokenType.EndOfStream && ConditionalCompilationContextStack.Count > 0 && !ConditionalCompilationContextStack.Peek().BlockActive)
                {
                    ResyncAt();
                    Read();

                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

        #endregion

        #region Public APIs
        private void StartParsing(Scanner scanner, IList<SyntaxError> errorTable)
        {
            StartParsing(scanner, errorTable, false, null, null, null, null);
        }

        private void StartParsing(Scanner scanner, IList<SyntaxError> errorTable, bool preprocess)
        {
            StartParsing(scanner, errorTable, preprocess, null, null, null, null);
        }

        private void StartParsing(Scanner scanner, IList<SyntaxError> errorTable, bool preprocess, IDictionary<string, object> conditionalCompilationConstants, IList<SourceRegion> sourceRegions, IList<ExternalLineMapping> externalLineMappings)
        {
            StartParsing(scanner, errorTable, preprocess, conditionalCompilationConstants, sourceRegions, externalLineMappings, null);
        }

        private void StartParsing(Scanner scanner, IList<SyntaxError> errorTable, bool preprocess, IDictionary<string, object> conditionalCompilationConstants, IList<SourceRegion> sourceRegions, IList<ExternalLineMapping> externalLineMappings, IList<ExternalChecksum> externalChecksums)
        {
            this.Scanner = scanner;
            this.ErrorTable = errorTable;
            this.Preprocess = preprocess;
            if (conditionalCompilationConstants == null)
            {
                this.ConditionalCompilationConstants = new Dictionary<string, object>();
            }
            else
            {
                // We have to clone this because the same hashtable could be used for
                // multiple parses.
                this.ConditionalCompilationConstants = new Dictionary<string, object>(conditionalCompilationConstants);
            }
            this.ExternalLineMappings = externalLineMappings;
            this.SourceRegions = sourceRegions;
            this.ExternalChecksums = externalChecksums;
            ErrorInConstruct = false;
            AtBeginningOfLine = true;
            BlockContextStack.Clear();
        }

        private void FinishParsing()
        {
            if (CurrentExternalSourceContext != null)
            {
                ReportSyntaxError(SyntaxErrorType.ExpectedEndExternalSource, Peek());
            }

            if (!(RegionContextStack.Count == 0))
            {
                ReportSyntaxError(SyntaxErrorType.ExpectedEndRegion, Peek());
            }

            if (!(ConditionalCompilationContextStack.Count == 0))
            {
                ReportSyntaxError(SyntaxErrorType.ExpectedCCEndIf, Peek());
            }

            StartParsing(null, null, false, null, null, null);
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Parse an entire file.
        /// </summary>
        /// <param name="scanner">The scanner to use to fetch the tokens.</param>
        /// <param name="errorTable">The list of errors produced during parsing.</param>
        /// <returns>A file-level parse tree.</returns>
        public FileTree ParseFile(Scanner scanner, IList<SyntaxError> errorTable)
        {
            FileTree File;

            StartParsing(scanner, errorTable, true);
            File = ParseFile();
            FinishParsing();

            return File;
        }

        /// <summary>
        /// Parse an entire file.
        /// </summary>
        /// <param name="scanner">The scanner to use to fetch the tokens.</param>
        /// <param name="errorTable">The list of errors produced during parsing.</param>
        /// <param name="conditionalCompilationConstants">Pre-defined conditional compilation constants.</param>
        /// <param name="sourceRegions">Source regions defined in the file.</param>
        /// <param name="externalLineMappings">External line mappings defined in the file.</param>
        /// <returns>A file-level parse tree.</returns>
        public FileTree ParseFile(Scanner scanner, IList<SyntaxError> errorTable, IDictionary<string, object> conditionalCompilationConstants, IList<SourceRegion> sourceRegions, IList<ExternalLineMapping> externalLineMappings, IList<ExternalChecksum> externalChecksums)
        {
            FileTree File;

            StartParsing(scanner, errorTable, true, conditionalCompilationConstants, sourceRegions, externalLineMappings, externalChecksums);
            File = ParseFile();
            FinishParsing();

            return File;
        }

        /// <summary>
        /// Parse a declaration.
        /// </summary>
        /// <param name="scanner">The scanner to use to fetch the tokens.</param>
        /// <param name="errorTable">The list of errors produced during parsing.</param>
        /// <returns>A declaration-level parse tree.</returns>
        public Declaration ParseDeclaration(Scanner scanner, IList<SyntaxError> errorTable)
        {
            Declaration Declaration;

            StartParsing(scanner, errorTable);
            Declaration = ParseDeclaration();
            FinishParsing();

            return Declaration;
        }

        /// <summary>
        /// Parse a statement.
        /// </summary>
        /// <param name="scanner">The scanner to use to fetch the tokens.</param>
        /// <param name="errorTable">The list of errors produced during parsing.</param>
        /// <returns>A statement-level parse tree.</returns>
        public Statement ParseStatement(Scanner scanner, IList<SyntaxError> errorTable)
        {
            Statement Statement;

            StartParsing(scanner, errorTable);
            Statement = ParseStatement();
            FinishParsing();

            return Statement;
        }

        /// <summary>
        /// Parse an expression.
        /// </summary>
        /// <param name="scanner">The scanner to use to fetch the tokens.</param>
        /// <param name="errorTable">The list of errors produced during parsing.</param>
        /// <returns>An expression-level parse tree.</returns>
        public Expression ParseExpression(Scanner scanner, IList<SyntaxError> errorTable)
        {
            Expression Expression;

            StartParsing(scanner, errorTable);
            Expression = ParseExpression();
            FinishParsing();

            return Expression;
        }

        /// <summary>
        /// Parse a type name.
        /// </summary>
        /// <param name="scanner">The scanner to use to fetch the tokens.</param>
        /// <param name="errorTable">The list of errors produced during parsing.</param>
        /// <returns>A typename-level parse tree.</returns>
        public TypeName ParseTypeName(Scanner scanner, IList<SyntaxError> errorTable)
        {
            TypeName TypeName;

            StartParsing(scanner, errorTable);
            TypeName = ParseTypeName(true);
            FinishParsing();

            return TypeName;
        }
        #endregion

    }
}
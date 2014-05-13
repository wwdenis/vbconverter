using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Xml;
using System.Text;

using Microsoft.VisualBasic;

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

namespace VBConverter.CodeParser.Serializers
{
    public sealed class TreeXmlSerializer : XmlSerializer<Tree>
    {
        #region Constuctor

        public TreeXmlSerializer(XmlWriter output) : base(output)
        {
        }

        public TreeXmlSerializer(XmlWriter output, bool outputPosition) : base(output, outputPosition)
        {
        }

        #endregion

        #region Static Methods

        private static TokenType GetOperatorToken(UnaryOperatorType type)
        {
            switch (type)
            {
                case UnaryOperatorType.Not:
                    return TokenType.Not;

                case UnaryOperatorType.UnaryPlus:
                    return TokenType.Plus;

                case UnaryOperatorType.Negate:
                    return TokenType.Minus;

                default:
                    return TokenType.LexicalError;
            }
        }

        private static TokenType GetOperatorToken(BinaryOperatorType type)
        {
            switch (type)
            {
                case BinaryOperatorType.Plus:
                    return TokenType.Plus;

                case BinaryOperatorType.Minus:
                    return TokenType.Minus;

                case BinaryOperatorType.Concatenate:
                    return TokenType.Ampersand;

                case BinaryOperatorType.Multiply:
                    return TokenType.Star;

                case BinaryOperatorType.Divide:
                    return TokenType.ForwardSlash;

                case BinaryOperatorType.IntegralDivide:
                    return TokenType.BackwardSlash;

                case BinaryOperatorType.Power:
                    return TokenType.Caret;

                case BinaryOperatorType.LessThan:
                    return TokenType.LessThan;

                case BinaryOperatorType.LessThanEquals:
                    return TokenType.LessThanEquals;

                case BinaryOperatorType.Equals:
                    return TokenType.Equals;

                case BinaryOperatorType.NotEquals:
                    return TokenType.NotEquals;

                case BinaryOperatorType.GreaterThan:
                    return TokenType.GreaterThan;

                case BinaryOperatorType.GreaterThanEquals:
                    return TokenType.GreaterThanEquals;

                case BinaryOperatorType.ShiftLeft:
                    return TokenType.LessThanLessThan;

                case BinaryOperatorType.ShiftRight:
                    return TokenType.GreaterThanGreaterThan;

                case BinaryOperatorType.Modulus:
                    return TokenType.Mod;

                case BinaryOperatorType.Or:
                    return TokenType.Or;

                case BinaryOperatorType.OrElse:
                    return TokenType.OrElse;

                case BinaryOperatorType.And:
                    return TokenType.And;

                case BinaryOperatorType.AndAlso:
                    return TokenType.AndAlso;

                case BinaryOperatorType.Xor:
                    return TokenType.Xor;

                case BinaryOperatorType.Like:
                    return TokenType.Like;

                case BinaryOperatorType.Is:
                    return TokenType.Is;

                case BinaryOperatorType.IsNot:
                    return TokenType.IsNot;

                case BinaryOperatorType.To:
                    return TokenType.To;

                default:
                    return TokenType.LexicalError;
            }
        }

        private static TokenType GetCompoundAssignmentOperatorToken(BinaryOperatorType compoundOperator)
        {
            switch (compoundOperator)
            {
                case BinaryOperatorType.Plus:
                    return TokenType.PlusEquals;

                case BinaryOperatorType.Concatenate:
                    return TokenType.AmpersandEquals;

                case BinaryOperatorType.Multiply:
                    return TokenType.StarEquals;

                case BinaryOperatorType.Minus:
                    return TokenType.MinusEquals;

                case BinaryOperatorType.Divide:
                    return TokenType.ForwardSlashEquals;

                case BinaryOperatorType.IntegralDivide:
                    return TokenType.BackwardSlashEquals;

                case BinaryOperatorType.Power:
                    return TokenType.CaretEquals;

                case BinaryOperatorType.ShiftLeft:
                    return TokenType.LessThanLessThanEquals;

                case BinaryOperatorType.ShiftRight:
                    return TokenType.GreaterThanGreaterThanEquals;

                default:
                    Debug.Fail("Unexpected!");
                    break;
            }

            return TokenType.None;
        }

        private static TokenType GetBlockTypeToken(BlockType blockType)
        {
            switch (blockType)
            {
                case BlockType.Class:
                    return TokenType.Class;

                case BlockType.Enum:
                    return TokenType.Enum;

                case BlockType.Function:
                    return TokenType.Function;

                case BlockType.Get:
                    return TokenType.Get;

                case BlockType.Event:
                    return TokenType.Event;

                case BlockType.AddHandler:
                    return TokenType.AddHandler;

                case BlockType.RemoveHandler:
                    return TokenType.RemoveHandler;

                case BlockType.RaiseEvent:
                    return TokenType.RaiseEvent;

                case BlockType.If:
                    return TokenType.If;

                case BlockType.Interface:
                    return TokenType.Interface;

                case BlockType.Module:
                    return TokenType.Module;

                case BlockType.Namespace:
                    return TokenType.Namespace;

                case BlockType.Property:
                    return TokenType.Property;

                case BlockType.Select:
                    return TokenType.Select;

                case BlockType.Set:
                    return TokenType.Set;

                case BlockType.Structure:
                    return TokenType.Structure;

                case BlockType.Sub:
                    return TokenType.Sub;

                case BlockType.SyncLock:
                    return TokenType.SyncLock;

                case BlockType.Using:
                    return TokenType.Using;

                case BlockType.Try:
                    return TokenType.Try;

                case BlockType.While:
                    return TokenType.While;

                case BlockType.With:
                    return TokenType.With;

                case BlockType.None:
                    return TokenType.LexicalError;

                case BlockType.Do:
                    return TokenType.Do;

                case BlockType.For:
                    return TokenType.For;

                case BlockType.Operator:
                    return TokenType.Operator;

                default:
                    Debug.Fail("Unexpected!");
                    break;
            }

            return TokenType.None;
        }

        #endregion

        #region Individual Trees

        private void SerializeToken(TokenType TokenType, Location CurrentLocation)
        {
            if (CurrentLocation.IsValid)
            {
                Output.WriteStartElement(TokenType.ToString());
                SerializePosition(CurrentLocation);
                Output.WriteEndElement();
            }
        }

        private void SerializeTypeCharacter(TypeCharacter TypeCharacter)
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

        private void SerializeArgument(Tree tree)
        {
            Serialize(((Argument)tree).Name);
            SerializeToken(TokenType.ByVal, ((Argument)tree).ByValLocation);
            SerializeToken(TokenType.ColonEquals, ((Argument)tree).ColonEqualsLocation);
            Serialize(((Argument)tree).Expression);
        }

        private void SerializeModifier(Tree tree)
        {
            Output.WriteAttributeString("type", ((Modifier)tree).ModifierType.ToString());
        }

        private void SerializeVariableDeclarator(Tree tree)
        {
            Output.WriteAttributeString("isField", ((VariableDeclarator)tree).IsField.ToString());

            Serialize(((VariableDeclarator)tree).VariableNames);
            SerializeToken(TokenType.As, ((VariableDeclarator)tree).AsLocation);
            SerializeToken(TokenType.New, ((VariableDeclarator)tree).NewLocation);
            Serialize(((VariableDeclarator)tree).VariableType);
            Serialize(((VariableDeclarator)tree).Arguments);
            SerializeToken(TokenType.Equals, ((VariableDeclarator)tree).EqualsLocation);
            Serialize(((VariableDeclarator)tree).Initializer);
        }

        private void SerializeTypeParameter(Tree tree)
        {
            Serialize(((TypeParameter)tree).TypeName);
            SerializeToken(TokenType.As, ((TypeParameter)tree).AsLocation);
            Serialize(((TypeParameter)tree).TypeConstraints);
        }

        private void SerializeParameter(Tree tree)
        {
            Serialize(((Parameter)tree).Attributes);
            Serialize(((Parameter)tree).Modifiers);
            Serialize(((Parameter)tree).VariableName);
            SerializeToken(TokenType.As, ((Parameter)tree).AsLocation);
            Serialize(((Parameter)tree).ParameterType);
            SerializeToken(TokenType.Equals, ((Parameter)tree).EqualsLocation);
            Serialize(((Parameter)tree).Initializer);
        }

        private void SerializeAttribute(Tree tree)
        {
            Output.WriteAttributeString("type", ((AttributeTree)tree).AttributeType.ToString());

            switch (((AttributeTree)tree).AttributeType)
            {
                case AttributeTypes.Assembly:
                    SerializeToken(TokenType.Colon, ((AttributeTree)tree).ColonLocation);
                    SerializeToken(TokenType.Assembly, ((AttributeTree)tree).AttributeTypeLocation);
                    break;

                case AttributeTypes.Module:
                    SerializeToken(TokenType.Module, ((AttributeTree)tree).AttributeTypeLocation);
                    SerializeToken(TokenType.Colon, ((AttributeTree)tree).ColonLocation);
                    break;
            }

            Serialize(((AttributeTree)tree).Name);
            Serialize(((AttributeTree)tree).Arguments);
        }

        #endregion

        #region Comments

        private void SerializeComments(Tree tree)
        {
            if (!(tree is CommentableTree))
                throw new ArgumentException("The provided tree is not a CommentableTree");

            if (((CommentableTree)tree).Comments != null)
                foreach (Comment comment in ((CommentableTree)tree).Comments)
                    Serialize(comment);
        }

        private void SerializeComment(Tree tree)
        {
            Output.WriteAttributeString("isRem", ((Comment)tree).IsREM.ToString());

            if ((((Comment)tree).Text != null))
                Output.WriteString(((Comment)tree).Text);
        }

        #endregion

        #region Collections

        private void SerializeCollection(Tree tree)
        {
            switch (tree.Type)
            {
                case TreeType.ArgumentCollection:
                    SerializeArgumentCollection(tree);
                    break;

                case TreeType.AttributeCollection:
                    SerializeAttributeCollection(tree);
                    break;

                case TreeType.CaseClauseCollection:
                    SerializeCaseClauseCollection(tree);
                    break;

                case TreeType.ExpressionCollection:
                    SerializeExpressionCollection(tree);
                    break;

                case TreeType.ImportCollection:
                    SerializeImportCollection(tree);
                    break;

                case TreeType.InitializerCollection:
                    SerializeInitializerCollection(tree);
                    break;

                case TreeType.NameCollection:
                    SerializeNameCollection(tree);
                    break;

                case TreeType.VariableNameCollection:
                    SerializeVariableNameCollection(tree);
                    break;

                case TreeType.ParameterCollection:
                    SerializeParameterCollection(tree);
                    break;

                case TreeType.TypeNameCollection:
                    SerializeTypeNameCollection(tree);
                    break;

                case TreeType.VariableDeclaratorCollection:
                    SerializeVariableDeclaratorCollection(tree);
                    break;

                case TreeType.DeclarationCollection:
                    SerializeDeclarationCollection(tree);
                    break;

                case TreeType.StatementCollection:
                    SerializeStatementCollection(tree);
                    break;

                case TreeType.TypeParameterCollection:
                    SerializeTypeParameterCollection(tree);
                    break;

                case TreeType.TypeConstraintCollection:
                    SerializeTypeConstraintCollection(tree);
                    break;

                case TreeType.TypeArgumentCollection:
                    SerializeTypeArgumentCollection(tree);
                    break;

                default:
                    foreach (Tree child in tree.Children)
                        Serialize(child);

                    break;
            }
        }

        private void SerializeArgumentCollection(Tree tree)
        {
            SerializeDelimitedList((ArgumentCollection)tree);

            if (((ArgumentCollection)tree).RightParenthesisLocation.IsValid)
                SerializeToken(TokenType.RightParenthesis, ((ArgumentCollection)tree).RightParenthesisLocation);
        }

        private void SerializeAttributeCollection(Tree tree)
        {
            SerializeDelimitedList((AttributeCollection)tree);

            if (((AttributeCollection)tree).RightBracketLocation.IsValid)
                SerializeToken(TokenType.GreaterThan, ((AttributeCollection)tree).RightBracketLocation);
        }

        private void SerializeCaseClauseCollection(Tree tree)
        {
            SerializeDelimitedList((CaseClauseCollection)tree);
        }

        private void SerializeExpressionCollection(Tree tree)
        {
            SerializeDelimitedList((ExpressionCollection)tree);
        }

        private void SerializeImportCollection(Tree tree)
        {
            SerializeDelimitedList((ImportCollection)tree);
        }

        private void SerializeInitializerCollection(Tree tree)
        {
            SerializeDelimitedList((InitializerCollection)tree);

            if (((InitializerCollection)tree).RightCurlyBraceLocation.IsValid)
                SerializeToken(TokenType.RightCurlyBrace, ((InitializerCollection)tree).RightCurlyBraceLocation);
        }

        private void SerializeNameCollection(Tree tree)
        {
            SerializeDelimitedList((NameCollection)tree);
        }

        private void SerializeVariableNameCollection(Tree tree)
        {
            SerializeDelimitedList((VariableNameCollection)tree);
        }

        private void SerializeParameterCollection(Tree tree)
        {
            SerializeDelimitedList((ParameterCollection)tree);

            if (((ParameterCollection)tree).RightParenthesisLocation.IsValid)
                SerializeToken(TokenType.RightParenthesis, ((ParameterCollection)tree).RightParenthesisLocation);
        }

        private void SerializeTypeNameCollection(Tree tree)
        {
            SerializeDelimitedList((TypeNameCollection)tree);
        }

        private void SerializeVariableDeclaratorCollection(Tree tree)
        {
            Output.WriteAttributeString("isField", ((VariableDeclaratorCollection)tree).IsField.ToString());
            SerializeDelimitedList((VariableDeclaratorCollection)tree);
        }

        private void SerializeDeclarationCollection(Tree tree)
        {
            SerializeDelimitedList((DeclarationCollection)tree);
        }

        private void SerializeStatementCollection(Tree tree)
        {
            SerializeDelimitedList((StatementCollection)tree);
        }

        private void SerializeTypeParameterCollection(Tree tree)
        {
            if (((TypeParameterCollection)tree).OfLocation.IsValid)
                SerializeToken(TokenType.Of, ((TypeParameterCollection)tree).OfLocation);

            SerializeDelimitedList((TypeParameterCollection)tree);

            if (((TypeParameterCollection)tree).RightParenthesisLocation.IsValid)
                SerializeToken(TokenType.RightParenthesis, ((TypeParameterCollection)tree).RightParenthesisLocation);
        }

        private void SerializeTypeConstraintCollection(Tree tree)
        {
            SerializeDelimitedList((TypeConstraintCollection)tree);

            if (((TypeConstraintCollection)tree).RightBracketLocation.IsValid)
                SerializeToken(TokenType.RightCurlyBrace, ((TypeConstraintCollection)tree).RightBracketLocation);
        }

        private void SerializeTypeArgumentCollection(Tree tree)
        {
            if (((TypeArgumentCollection)tree).OfLocation.IsValid)
                SerializeToken(TokenType.Of, ((TypeArgumentCollection)tree).OfLocation);

            SerializeDelimitedList((TypeArgumentCollection)tree);

            if (((TypeArgumentCollection)tree).RightParenthesisLocation.IsValid)
                SerializeToken(TokenType.RightParenthesis, ((TypeArgumentCollection)tree).RightParenthesisLocation);
        }

        private void SerializeDelimitedList<T>(ColonDelimitedTreeCollection<T> list) where T : Tree
        {
            IEnumerator<Location> ColonEnumerator;
            bool MoreColons;

            if ((list.ColonLocations != null))
            {
                ColonEnumerator = list.ColonLocations.GetEnumerator();
                MoreColons = ColonEnumerator.MoveNext();
            }
            else
            {
                ColonEnumerator = null;
                MoreColons = false;
            }

            foreach (Tree Child in list.Children)
            {
                while (MoreColons && ColonEnumerator.Current <= Child.Span.Start)
                {
                    SerializeToken(TokenType.Colon, ColonEnumerator.Current);
                    MoreColons = ColonEnumerator.MoveNext();
                }

                Serialize(Child);
            }

            while (MoreColons)
            {
                SerializeToken(TokenType.Colon, ColonEnumerator.Current);
                MoreColons = ColonEnumerator.MoveNext();
            }
        }

        private void SerializeDelimitedList<T>(CommaDelimitedTreeCollection<T> list) where T : Tree
        {
            IEnumerator<Location> CommaEnumerator;
            bool MoreCommas;

            if ((list.CommaLocations != null))
            {
                CommaEnumerator = list.CommaLocations.GetEnumerator();
                MoreCommas = CommaEnumerator.MoveNext();
            }
            else
            {
                CommaEnumerator = null;
                MoreCommas = false;
            }

            foreach (Tree Child in list.Children)
            {
                if ((Child != null))
                {
                    while (MoreCommas && CommaEnumerator.Current <= Child.Span.Start)
                    {
                        SerializeToken(TokenType.Comma, CommaEnumerator.Current);
                        MoreCommas = CommaEnumerator.MoveNext();
                    }

                    Serialize(Child);
                }
            }

            while (MoreCommas)
            {
                SerializeToken(TokenType.Comma, CommaEnumerator.Current);
                MoreCommas = CommaEnumerator.MoveNext();
            }
        }

        #endregion

        #region Names

        private void SerializeName(Tree Name)
        {
            switch (Name.Type)
            {
                case TreeType.SimpleName:
                    SerializeSimpleName(Name);
                    break;

                case TreeType.VariableName:
                    SerializeVariableName(Name);
                    break;

                case TreeType.QualifiedName:
                    SerializeQualifiedName(Name);
                    break;
            }
        }

        private void SerializeSimpleName(Tree Name)
        {
            SerializeTypeCharacter(((SimpleName)Name).TypeCharacter);
            Output.WriteAttributeString("escaped", ((SimpleName)Name).Escaped.ToString());
            Output.WriteString(((SimpleName)Name).Name);
        }

        private void SerializeVariableName(Tree Name)
        {
            Serialize(((VariableName)Name).Name);
            Serialize(((VariableName)Name).ArrayType);
        }

        private void SerializeQualifiedName(Tree Name)
        {
            Serialize(((QualifiedName)Name).Qualifier);
            SerializeToken(TokenType.Period, ((QualifiedName)Name).DotLocation);
            Serialize(((QualifiedName)Name).Name);
        }

        #endregion

        #region Types

        private void SerializeTypeName(Tree Type)
        {
            switch (Type.Type)
            {
                case TreeType.IntrinsicType:
                    SerializeIntrinsicTypeName(Type);
                    break;

                case TreeType.NamedType:
                    SerializeNamedTypeName(Type);
                    break;

                case TreeType.ArrayType:
                    SerializeArrayTypeName(Type);
                    break;

                case TreeType.ConstructedType:
                    SerializeConstructedTypeName(Type);
                    break;
            }
        }

        private void SerializeIntrinsicTypeName(Tree Type)
        {
            Output.WriteAttributeString("intrinsicType", ((IntrinsicTypeName)Type).IntrinsicType.ToString());
            Serialize(((IntrinsicTypeName)Type).StringLength);
        }

        private void SerializeNamedTypeName(Tree Type)
        {
            Serialize(((NamedTypeName)Type).Name);
        }

        private void SerializeArrayTypeName(Tree Type)
        {
            Output.WriteAttributeString("rank", ((ArrayTypeName)Type).Rank.ToString());
            Serialize(((ArrayTypeName)Type).ElementTypeName);
            Serialize(((ArrayTypeName)Type).Arguments);
        }

        private void SerializeConstructedTypeName(Tree Type)
        {
            Serialize(((ConstructedTypeName)Type).Name);
            Serialize(((ConstructedTypeName)Type).TypeArguments);
        }

        #endregion

        #region Initializers

        private void SerializeInitializer(Tree Initializer)
        {
            switch (Initializer.Type)
            {
                case TreeType.AggregateInitializer:
                    SerializeAggregateInitializer(Initializer);
                    break;

                case TreeType.ExpressionInitializer:
                    SerializeExpressionInitializer(Initializer);
                    break;
            }
        }

        private void SerializeAggregateInitializer(Tree Initializer)
        {
            Serialize(((AggregateInitializer)Initializer).Elements);
        }

        private void SerializeExpressionInitializer(Tree Initializer)
        {
            Serialize(((ExpressionInitializer)Initializer).Expression);
        }

        #endregion

        #region Imports

        private void SerializeImport(Tree Import)
        {
            switch (Import.Type)
            {
                case TreeType.NameImport:
                    SerializeNameImport(Import);
                    break;

                case TreeType.AliasImport:
                    SerializeAliasImport(Import);
                    break;
            }
        }

        private void SerializeAliasImport(Tree Import)
        {
            Serialize(((AliasImport)Import).AliasedTypeName);
            SerializeToken(TokenType.Equals, ((AliasImport)Import).EqualsLocation);
            Serialize(((AliasImport)Import).Name);
        }

        private void SerializeNameImport(Tree Import)
        {
            Serialize(((NameImport)Import).TypeName);
        }

        #endregion

        #region Case Clause

        private void SerializeCaseClause(Tree CaseClause)
        {
            switch (CaseClause.Type)
            {
                case TreeType.ComparisonCaseClause:
                    SerializeComparisonCaseClause(CaseClause);
                    break;

                case TreeType.RangeCaseClause:
                    SerializeRangeCaseClause(CaseClause);
                    break;

                default:
                    Debug.Fail("Unexpected.");
                    break;
            }
        }

        private void SerializeComparisonCaseClause(Tree CaseClause)
        {
            SerializeToken(TokenType.Is, ((ComparisonCaseClause)CaseClause).IsLocation);
            SerializeToken(GetOperatorToken(((ComparisonCaseClause)CaseClause).ComparisonOperator), ((ComparisonCaseClause)CaseClause).OperatorLocation);
            Serialize(((ComparisonCaseClause)CaseClause).Operand);
        }

        private void SerializeRangeCaseClause(Tree CaseClause)
        {
            Serialize(((RangeCaseClause)CaseClause).RangeExpression);
        }

        #endregion

        #region Expressions

        private void SerializeExpression(Tree expression)
        {
            switch (expression.Type)
            {
                case TreeType.CallOrIndexExpression:
                    SerializeCallOrIndexExpression(expression);
                    break;

                case TreeType.BooleanLiteralExpression:
                case TreeType.StringLiteralExpression:
                case TreeType.CharacterLiteralExpression:
                case TreeType.DateLiteralExpression:
                case TreeType.IntegerLiteralExpression:
                case TreeType.FloatingPointLiteralExpression:
                case TreeType.DecimalLiteralExpression:
                    SerializeLiteralExpression(expression);
                    break;

                case TreeType.GetTypeExpression:
                    SerializeGetTypeExpression(expression);
                    break;

                case TreeType.CTypeExpression:
                case TreeType.DirectCastExpression:
                    SerializeCastTypeExpression(expression);
                    break;

                case TreeType.TypeOfExpression:
                    SerializeTypeOfExpression(expression);
                    break;

                case TreeType.IntrinsicCastExpression:
                    SerializeIntrinsicCastExpression(expression);
                    break;

                case TreeType.SimpleNameExpression:
                    SerializeSimpleNameExpression(expression);
                    break;

                case TreeType.QualifiedExpression:
                    SerializeQualifiedExpression(expression);
                    break;

                case TreeType.DictionaryLookupExpression:
                    SerializeDictionaryLookupExpression(expression);
                    break;

                case TreeType.InstanceExpression:
                    SerializeInstanceExpression(expression);
                    break;

                case TreeType.ParentheticalExpression:
                    SerializeParentheticalExpression(expression);
                    break;

                case TreeType.BinaryOperatorExpression:
                    SerializeBinaryOperatorExpression(expression);
                    break;

                case TreeType.UnaryOperatorExpression:
                    SerializeUnaryOperatorExpression(expression);
                    break;

                case TreeType.NewExpression:
                    SerializeNewExpression(expression);
                    break;

                case TreeType.NothingExpression:
                    SerializeNothingExpression(expression);
                    break;

                default:
                    foreach (Tree child in expression.Children)
                        Serialize(child);
                    break;
            }
        }

        private void SerializeCallOrIndexExpression(Tree expression)
        {
            Output.WriteAttributeString("isIndex", ((CallOrIndexExpression)expression).IsIndex.ToString());
            foreach (Tree child in expression.Children)
                Serialize(child);
        }

        private void SerializeLiteralExpression(Tree expression)
        {
            Type treeType = expression.GetType();

            Output.WriteAttributeString("isConstant", ((Expression)expression).IsConstant.ToString());

            switch (expression.Type)
            {
                case TreeType.BooleanLiteralExpression:
                    Output.WriteString(((BooleanLiteralExpression)expression).Literal.ToString());
                    break;

                case TreeType.StringLiteralExpression:
                    Output.WriteString(((StringLiteralExpression)expression).Literal.ToString());
                    break;

                case TreeType.CharacterLiteralExpression:
                    Output.WriteString(((CharacterLiteralExpression)expression).Literal.ToString());
                    break;

                case TreeType.DateLiteralExpression:
                    Output.WriteString(((DateLiteralExpression)expression).Literal.ToString());
                    break;

                case TreeType.IntegerLiteralExpression:
                    SerializeTypeCharacter(((IntegerLiteralExpression)expression).TypeCharacter);
                    Output.WriteAttributeString("base", ((IntegerLiteralExpression)expression).IntegerBase.ToString());
                    Output.WriteString(((IntegerLiteralExpression)expression).Literal.ToString());
                    break;

                case TreeType.FloatingPointLiteralExpression:
                    SerializeTypeCharacter(((FloatingPointLiteralExpression)expression).TypeCharacter);
                    Output.WriteString(((FloatingPointLiteralExpression)expression).Literal.ToString());
                    break;

                case TreeType.DecimalLiteralExpression:
                    SerializeTypeCharacter(((DecimalLiteralExpression)expression).TypeCharacter);
                    Output.WriteString(((DecimalLiteralExpression)expression).Literal.ToString());
                    break;
            }
        }

        private void SerializeGetTypeExpression(Tree expression)
        {
            Output.WriteAttributeString("isConstant", ((Expression)expression).IsConstant.ToString());
            SerializeToken(TokenType.LeftParenthesis, ((GetTypeExpression)expression).LeftParenthesisLocation);
            Serialize(((GetTypeExpression)expression).Target);
            SerializeToken(TokenType.RightParenthesis, ((GetTypeExpression)expression).RightParenthesisLocation);
        }

        private void SerializeCastTypeExpression(Tree expression)
        {
            Output.WriteAttributeString("isConstant", ((Expression)expression).IsConstant.ToString());
            SerializeToken(TokenType.LeftParenthesis, ((CastTypeExpression)expression).LeftParenthesisLocation);
            Serialize(((CastTypeExpression)expression).Operand);
            SerializeToken(TokenType.Comma, ((CastTypeExpression)expression).CommaLocation);
            Serialize(((CastTypeExpression)expression).Target);
            SerializeToken(TokenType.RightParenthesis, ((CastTypeExpression)expression).RightParenthesisLocation);
        }

        private void SerializeTypeOfExpression(Tree expression)
        {
            Output.WriteAttributeString("isConstant", ((Expression)expression).IsConstant.ToString());
            Serialize(((TypeOfExpression)expression).Operand);
            SerializeToken(TokenType.Is, ((TypeOfExpression)expression).IsLocation);
            Serialize(((TypeOfExpression)expression).Target);
        }

        private void SerializeIntrinsicCastExpression(Tree expression)
        {
            Output.WriteAttributeString("isConstant", ((Expression)expression).IsConstant.ToString());
            Output.WriteAttributeString("intrinsicType", ((IntrinsicCastExpression)expression).IntrinsicType.ToString());
            SerializeToken(TokenType.LeftParenthesis, ((IntrinsicCastExpression)expression).LeftParenthesisLocation);
            Serialize(((IntrinsicCastExpression)expression).Operand);
            SerializeToken(TokenType.RightParenthesis, ((IntrinsicCastExpression)expression).RightParenthesisLocation);
        }

        private void SerializeSimpleNameExpression(Tree expression)
        {
            foreach (Tree child in expression.Children)
                Serialize(child);
        }

        private void SerializeQualifiedExpression(Tree expression)
        {
            Output.WriteAttributeString("isConstant", ((Expression)expression).IsConstant.ToString());
            Serialize(((QualifiedExpression)expression).Qualifier);
            SerializeToken(TokenType.Period, ((QualifiedExpression)expression).DotLocation);
            Serialize(((QualifiedExpression)expression).Name);
        }

        private void SerializeDictionaryLookupExpression(Tree expression)
        {
            Output.WriteAttributeString("isConstant", ((Expression)expression).IsConstant.ToString());
            Serialize(((DictionaryLookupExpression)expression).Qualifier);
            SerializeToken(TokenType.Exclamation, ((DictionaryLookupExpression)expression).BangLocation);
            Serialize(((DictionaryLookupExpression)expression).Name);
        }

        private void SerializeInstanceExpression(Tree expression)
        {
            Output.WriteAttributeString("isConstant", ((Expression)expression).IsConstant.ToString());
            Output.WriteAttributeString("type", ((InstanceExpression)expression).InstanceType.ToString());
        }

        private void SerializeParentheticalExpression(Tree expression)
        {
            Output.WriteAttributeString("isConstant", ((Expression)expression).IsConstant.ToString());
            Serialize(((ParentheticalExpression)expression).Operand);
            SerializeToken(TokenType.RightParenthesis, ((ParentheticalExpression)expression).RightParenthesisLocation);
        }

        private void SerializeBinaryOperatorExpression(Tree expression)
        {
            Output.WriteAttributeString("isConstant", ((Expression)expression).IsConstant.ToString());
            Output.WriteAttributeString("operator", ((BinaryOperatorExpression)expression).Operator.ToString());
            Serialize(((BinaryOperatorExpression)expression).LeftOperand);
            SerializeToken(GetOperatorToken(((BinaryOperatorExpression)expression).Operator), ((BinaryOperatorExpression)expression).OperatorLocation);
            Serialize(((BinaryOperatorExpression)expression).RightOperand);
        }


        private void SerializeUnaryOperatorExpression(Tree expression)
        {
            Output.WriteAttributeString("isConstant", ((Expression)expression).IsConstant.ToString());
            SerializeToken(GetOperatorToken(((UnaryOperatorExpression)expression).Operator), ((UnaryOperatorExpression)expression).Span.Start);
            Serialize(((UnaryOperatorExpression)expression).Operand);
        }

        private void SerializeNewExpression(Tree expression)
        {
            foreach (Tree child in expression.Children)
                Serialize(child);
        }

        private void SerializeNothingExpression(Tree expression)
        {
            foreach (Tree child in expression.Children)
                Serialize(child);
        }

        #endregion

        #region Statements

        private void SerializeStatement(Tree statement)
        {
            switch (statement.Type)
            {
                case TreeType.GotoStatement:
                    SerializeGotoStatement(statement);
                    break;

                case TreeType.GoSubStatement:
                    SerializeGoSubStatement(statement);
                    break;

                case TreeType.LabelStatement:
                    SerializeLabelStatement(statement);
                    break;

                case TreeType.ContinueStatement:
                    SerializeContinueStatement(statement);
                    break;

                case TreeType.ExitStatement:
                    SerializeExitStatement(statement);
                    break;

                case TreeType.ReturnStatement:
                    SerializeReturnStatement(statement);
                    break;

                case TreeType.ErrorStatement:
                    SerializeErrorStatement(statement);
                    break;

                case TreeType.ThrowStatement:
                    SerializeThrowStatement(statement);
                    break;

                case TreeType.RaiseEventStatement:
                    SerializeRaiseEventStatement(statement);
                    break;

                case TreeType.AddHandlerStatement:
                case TreeType.RemoveHandlerStatement:
                    SerializeHandlerStatement(statement);
                    break;

                case TreeType.OnErrorStatement:
                    SerializeOnErrorStatement(statement);
                    break;

                case TreeType.ResumeStatement:
                    SerializeResumeStatement(statement);
                    break;

                case TreeType.ReDimStatement:
                    SerializeReDimStatement(statement);
                    break;

                case TreeType.EraseStatement:
                    SerializeEraseStatement(statement);
                    break;

                case TreeType.CallStatement:
                    SerializeCallStatement(statement);
                    break;

                case TreeType.AssignmentStatement:
                    SerializeAssignmentStatement(statement);
                    break;

                case TreeType.MidAssignmentStatement:
                    SerializeMidAssignmentStatement(statement);
                    break;

                case TreeType.CompoundAssignmentStatement:
                    SerializeCompoundAssignmentStatement(statement);
                    break;

                case TreeType.LocalDeclarationStatement:
                    SerializeLocalDeclarationStatement(statement);
                    break;

                case TreeType.EndBlockStatement:
                    SerializeEndBlockStatement(statement);
                    break;

                case TreeType.WhileBlockStatement:
                    SerializeWhileBlockStatement(statement);
                    break;

                case TreeType.WithBlockStatement:
                    SerializeWithBlockStatement(statement);
                    break;

                case TreeType.SyncLockBlockStatement:
                    SerializeSyncLockBlockStatement(statement);
                    break;

                case TreeType.UsingBlockStatement:
                    SerializeComments(statement);
                    SerializeUsingBlockStatement(statement);
                    break;

                case TreeType.DoBlockStatement:
                    SerializeDoBlockStatement(statement);
                    break;

                case TreeType.LoopStatement:
                    SerializeLoopStatement(statement);
                    break;

                case TreeType.NextStatement:
                    SerializeNextStatement(statement);
                    break;

                case TreeType.ForBlockStatement:
                    SerializeForBlockStatement(statement);
                    break;

                case TreeType.ForEachBlockStatement:
                    SerializeForEachBlockStatement(statement);
                    break;

                case TreeType.CatchStatement:
                    SerializeCatchStatement(statement);
                    break;

                case TreeType.CaseStatement:
                    SerializeCaseStatement(statement);
                    break;

                case TreeType.CaseElseStatement:
                    SerializeCaseElseStatement(statement);
                    break;

                case TreeType.CaseBlockStatement:
                    SerializeCaseBlockStatement(statement);
                    break;

                case TreeType.CaseElseBlockStatement:
                    SerializeCaseElseBlockStatement(statement);
                    break;   
                
                case TreeType.SelectBlockStatement:
                    SerializeSelectBlockStatement(statement);
                    break;

                case TreeType.IfBlockStatement:
                    SerializeIfBlockStatement(statement);
                    break;

                case TreeType.ElseBlockStatement:
                    SerializeElseBlockStatement(statement);
                    break;

                case TreeType.ElseIfStatement:
                    SerializeElseIfStatement(statement);
                    break;

                case TreeType.ElseIfBlockStatement:
                    SerializeElseIfBlockStatement(statement);
                    break;

                case TreeType.LineIfBlockStatement:
                    SerializeLineIfBlockStatement(statement);
                    break;

                case TreeType.EmptyStatement:
                    SerializeEmptyStatement(statement);
                    break;

                case TreeType.EndStatement:
                    SerializeEndStatement(statement);
                    break;

                default:
                    SerializeComments(statement);
                    foreach (Tree child in statement.Children)
                        Serialize(child);
                    break;
            }
        }

        private void SerializeLabelStatement(Tree statement)
        {
            Output.WriteAttributeString("isLineNumber", ((LabelStatement)statement).IsLineNumber.ToString());
            SerializeComments(statement);
            Serialize(((LabelStatement)statement).Name);
        }

        private void SerializeGotoStatement(Tree statement)
        {
            Output.WriteAttributeString("isLineNumber", ((GotoStatement)statement).IsLineNumber.ToString());
            SerializeComments(statement);
            Serialize(((GotoStatement)statement).Name);
        }

        private void SerializeGoSubStatement(Tree statement)
        {
            Output.WriteAttributeString("isLineNumber", ((GoSubStatement)statement).IsLineNumber.ToString());
            SerializeComments(statement);
            Serialize(((GoSubStatement)statement).Name);
        }

        private void SerializeContinueStatement(Tree statement)
        {
            Output.WriteAttributeString("continueType", ((ContinueStatement)statement).ContinueType.ToString());
            SerializeComments(statement);
            SerializeToken(GetBlockTypeToken(((ContinueStatement)statement).ContinueType), ((ContinueStatement)statement).ContinueArgumentLocation);
        }

        private void SerializeExitStatement(Tree statement)
        {
            Output.WriteAttributeString("exitType", ((ExitStatement)statement).ExitType.ToString());
            
            if (((ExitStatement)statement).RelatedTree != null)
                Output.WriteAttributeString("relatedTreeType", ((ExitStatement)statement).RelatedTree.Type.ToString());

            SerializeComments(statement);
            SerializeToken(GetBlockTypeToken(((ExitStatement)statement).ExitType), ((ExitStatement)statement).ExitArgumentLocation);
        }

        private void SerializeReturnStatement(Tree statement)
        {
            SerializeComments(statement);
            Serialize(((ReturnStatement)statement).Expression);
            Serialize(((ReturnStatement)statement).GoSubReference);
        }

        private void SerializeErrorStatement(Tree statement)
        {
            SerializeComments(statement);
            Serialize(((ErrorStatement)statement).Expression);
        }

        private void SerializeThrowStatement(Tree statement)
        {
            SerializeComments(statement);
            Serialize(((ThrowStatement)statement).Expression);
        }

        private void SerializeRaiseEventStatement(Tree statement)
        {
            SerializeComments(statement);
            Serialize(((RaiseEventStatement)statement).Name);
            Serialize(((RaiseEventStatement)statement).Arguments);
        }

        private void SerializeHandlerStatement(Tree statement)
        {
            SerializeComments(statement);
            Serialize(((HandlerStatement)statement).Name);
            SerializeToken(TokenType.Comma, ((HandlerStatement)statement).CommaLocation);
            Serialize(((HandlerStatement)statement).DelegateExpression);
        }

        private void SerializeOnErrorStatement(Tree statement)
        {
            Output.WriteAttributeString("onErrorType", ((OnErrorStatement)statement).OnErrorType.ToString());
            SerializeComments(statement);
            SerializeToken(TokenType.Error, ((OnErrorStatement)statement).ErrorLocation);

            switch (((OnErrorStatement)statement).OnErrorType)
            {
                case OnErrorType.Zero:
                    SerializeToken(TokenType.GoTo, ((OnErrorStatement)statement).ResumeOrGoToLocation);
                    Output.WriteStartElement("Zero");
                    SerializePosition(((OnErrorStatement)statement).NextOrZeroOrMinusLocation);
                    Output.WriteEndElement();
                    break;

                case OnErrorType.MinusOne:
                    SerializeToken(TokenType.GoTo, ((OnErrorStatement)statement).ResumeOrGoToLocation);
                    SerializeToken(TokenType.Minus, ((OnErrorStatement)statement).NextOrZeroOrMinusLocation);
                    Output.WriteStartElement("One");
                    SerializePosition(((OnErrorStatement)statement).OneLocation);
                    Output.WriteEndElement();
                    break;

                case OnErrorType.Label:
                    SerializeToken(TokenType.GoTo, ((OnErrorStatement)statement).ResumeOrGoToLocation);
                    Serialize(((OnErrorStatement)statement).Name);
                    break;

                case OnErrorType.Next:
                    SerializeToken(TokenType.Resume, ((OnErrorStatement)statement).ResumeOrGoToLocation);
                    SerializeToken(TokenType.Next, ((OnErrorStatement)statement).NextOrZeroOrMinusLocation);
                    break;

                case OnErrorType.Bad:
                    break;
                // Do nothing
            }
        }

        private void SerializeResumeStatement(Tree statement)
        {
            Output.WriteAttributeString("resumeType", ((ResumeStatement)statement).ResumeType.ToString());
            SerializeComments(statement);
            switch (((ResumeStatement)statement).ResumeType)
            {
                case ResumeType.Next:
                    SerializeToken(TokenType.Next, ((ResumeStatement)statement).NextLocation);
                    break;

                case ResumeType.Label:
                    Serialize(((ResumeStatement)statement).Name);
                    break;

                case ResumeType.None:
                    break;
                // Do nothing
            }
        }

        private void SerializeReDimStatement(Tree statement)
        {
            SerializeComments(statement);
            SerializeToken(TokenType.Preserve, ((ReDimStatement)statement).PreserveLocation);
            Serialize(((ReDimStatement)statement).Variables);
        }

        private void SerializeEraseStatement(Tree statement)
        {
            SerializeComments(statement);
            Serialize(((EraseStatement)statement).Variables);
        }

        private void SerializeCallStatement(Tree statement)
        {
            SerializeComments(statement);
            SerializeToken(TokenType.Call, ((CallStatement)statement).CallLocation);
            Serialize(((CallStatement)statement).TargetExpression);
            Serialize(((CallStatement)statement).Arguments);
        }

        private void SerializeAssignmentStatement(Tree statement)
        {
            AssignmentStatement Current = (AssignmentStatement)statement;

            SerializeComments(statement);
            SerializeToken(Current.AcessorType, Current.AcessorLocation);
            Serialize(Current.TargetExpression);
            SerializeToken(TokenType.Equals, Current.OperatorLocation);
            Serialize(Current.SourceExpression);
        }

        private void SerializeMidAssignmentStatement(Tree statement)
        {
            Output.WriteAttributeString("hasTypeCharacter", ((MidAssignmentStatement)statement).HasTypeCharacter.ToString());
            SerializeComments(statement);
            SerializeToken(TokenType.LeftParenthesis, ((MidAssignmentStatement)statement).LeftParenthesisLocation);
            Serialize(((MidAssignmentStatement)statement).TargetExpression);
            SerializeToken(TokenType.Comma, ((MidAssignmentStatement)statement).StartCommaLocation);
            Serialize(((MidAssignmentStatement)statement).StartExpression);
            SerializeToken(TokenType.Comma, ((MidAssignmentStatement)statement).LengthCommaLocation);
            Serialize(((MidAssignmentStatement)statement).LengthExpression);
            SerializeToken(TokenType.RightParenthesis, ((MidAssignmentStatement)statement).RightParenthesisLocation);
            SerializeToken(TokenType.Equals, ((MidAssignmentStatement)statement).OperatorLocation);
            Serialize(((MidAssignmentStatement)statement).SourceExpression);
        }

        private void SerializeCompoundAssignmentStatement(Tree statement)
        {
            SerializeComments(statement);
            Serialize(((CompoundAssignmentStatement)statement).TargetExpression);
            SerializeToken(GetCompoundAssignmentOperatorToken(((CompoundAssignmentStatement)statement).CompoundOperator), ((CompoundAssignmentStatement)statement).OperatorLocation);
            Serialize(((CompoundAssignmentStatement)statement).SourceExpression);
        }

        private void SerializeLocalDeclarationStatement(Tree statement)
        {
            SerializeComments(statement);
            Serialize(((LocalDeclarationStatement)statement).Modifiers);
            Serialize(((LocalDeclarationStatement)statement).VariableDeclarators);
        }

        private void SerializeEndBlockStatement(Tree statement)
        {
            Output.WriteAttributeString("endType", ((EndBlockStatement)statement).EndType.ToString());
            SerializeComments(statement);
            SerializeToken(GetBlockTypeToken(((EndBlockStatement)statement).EndType), ((EndBlockStatement)statement).EndArgumentLocation);
        }

        private void SerializeWhileBlockStatement(Tree statement)
        {
            SerializeComments(statement);
            Serialize(((WhileBlockStatement)statement).Expression);
            Serialize(((WhileBlockStatement)statement).Statements);
            Serialize(((WhileBlockStatement)statement).EndStatement);
        }

        private void SerializeWithBlockStatement(Tree statement)
        {
            SerializeComments(statement);
            Serialize(((WithBlockStatement)statement).Expression);
            Serialize(((WithBlockStatement)statement).Statements);
            Serialize(((WithBlockStatement)statement).EndStatement);
        }

        private void SerializeSyncLockBlockStatement(Tree statement)
        {
            SerializeComments(statement);
            Serialize(((SyncLockBlockStatement)statement).Expression);
            Serialize(((SyncLockBlockStatement)statement).Statements);
            Serialize(((SyncLockBlockStatement)statement).EndStatement);
        }

        private void SerializeUsingBlockStatement(Tree statement)
        {
            if (((UsingBlockStatement)statement).Expression != null)
                Serialize(((UsingBlockStatement)statement).Expression);
            else
                Serialize(((UsingBlockStatement)statement).VariableDeclarators);
            Serialize(((UsingBlockStatement)statement).Statements);
            Serialize(((UsingBlockStatement)statement).EndStatement);
        }

        private void SerializeDoBlockStatement(Tree statement)
        {
            if ((((DoBlockStatement)statement).Expression != null))
            {
                Output.WriteAttributeString("isWhile", Conversion.Str(((DoBlockStatement)statement).IsWhile));
                SerializeComments(statement);
                if (((DoBlockStatement)statement).IsWhile)
                    SerializeToken(TokenType.While, ((DoBlockStatement)statement).WhileOrUntilLocation);
                else
                    SerializeToken(TokenType.Until, ((DoBlockStatement)statement).WhileOrUntilLocation);
                Serialize(((DoBlockStatement)statement).Expression);
            }
            else
            {
                SerializeComments(statement);
            }
            Serialize(((DoBlockStatement)statement).Statements);
            Serialize(((DoBlockStatement)statement).EndStatement);
        }

        private void SerializeLoopStatement(Tree statement)
        {
            if ((((LoopStatement)statement).Expression != null))
            {
                Output.WriteAttributeString("isWhile", Conversion.Str(((LoopStatement)statement).IsWhile));
                SerializeComments(statement);
                if (((LoopStatement)statement).IsWhile)
                    SerializeToken(TokenType.While, ((LoopStatement)statement).WhileOrUntilLocation);
                else
                    SerializeToken(TokenType.Until, ((LoopStatement)statement).WhileOrUntilLocation);
                Serialize(((LoopStatement)statement).Expression);
            }
            else
            {
                SerializeComments(statement);
            }
        }

        private void SerializeNextStatement(Tree statement)
        {
            SerializeComments(statement);
            Serialize(((NextStatement)statement).Variables);
        }

        private void SerializeForBlockStatement(Tree statement)
        {
            SerializeComments(statement);
            Serialize(((ForBlockStatement)statement).ControlExpression);
            Serialize(((ForBlockStatement)statement).ControlVariableDeclarator);
            SerializeToken(TokenType.Equals, ((ForBlockStatement)statement).EqualsLocation);
            Serialize(((ForBlockStatement)statement).LowerBoundExpression);
            SerializeToken(TokenType.To, ((ForBlockStatement)statement).ToLocation);
            Serialize(((ForBlockStatement)statement).UpperBoundExpression);
            SerializeToken(TokenType.Step, ((ForBlockStatement)statement).StepLocation);
            Serialize(((ForBlockStatement)statement).StepExpression);
            Serialize(((ForBlockStatement)statement).Statements);
            Serialize(((ForBlockStatement)statement).NextStatement);
        }

        private void SerializeForEachBlockStatement(Tree statement)
        {
            SerializeComments(statement);
            SerializeToken(TokenType.Each, ((ForEachBlockStatement)statement).EachLocation);
            Serialize(((ForEachBlockStatement)statement).ControlExpression);
            Serialize(((ForEachBlockStatement)statement).ControlVariableDeclarator);
            SerializeToken(TokenType.In, ((ForEachBlockStatement)statement).InLocation);
            Serialize(((ForEachBlockStatement)statement).CollectionExpression);
            Serialize(((ForEachBlockStatement)statement).Statements);
            Serialize(((ForEachBlockStatement)statement).NextStatement);
        }

        private void SerializeCatchStatement(Tree statement)
        {
            SerializeComments(statement);
            Serialize(((CatchStatement)statement).Name);
            SerializeToken(TokenType.As, ((CatchStatement)statement).AsLocation);
            Serialize(((CatchStatement)statement).ExceptionType);
            SerializeToken(TokenType.When, ((CatchStatement)statement).WhenLocation);
            Serialize(((CatchStatement)statement).FilterExpression);
        }

        private void SerializeCaseStatement(Tree statement)
        {
            SerializeComments(statement);
            foreach (Tree child in statement.Children)
                Serialize(child);
        }

        private void SerializeCaseElseStatement(Tree statement)
        {
            SerializeComments(statement);
            SerializeToken(TokenType.Else, ((CaseElseStatement)statement).ElseLocation);
        }

        private void SerializeCaseBlockStatement(Tree statement)
        {
            SerializeComments(statement);
            foreach (Tree child in statement.Children)
                Serialize(child);
        }

        private void SerializeCaseElseBlockStatement(Tree statement)
        {
            SerializeComments(statement);
            foreach (Tree child in statement.Children)
                Serialize(child);
        }

        private void SerializeSelectBlockStatement(Tree statement)
        {
            SerializeComments(statement);
            SerializeToken(TokenType.Case, ((SelectBlockStatement)statement).CaseLocation);
            Serialize(((SelectBlockStatement)statement).Expression);
            Serialize(((SelectBlockStatement)statement).Statements);
            Serialize(((SelectBlockStatement)statement).CaseBlockStatements);
            Serialize(((SelectBlockStatement)statement).CaseElseBlockStatement);
            Serialize(((SelectBlockStatement)statement).EndStatement);
        }

        private void SerializeIfBlockStatement(Tree statement)
        {
            SerializeComments(statement);
            Serialize(((IfBlockStatement)statement).Expression);
            SerializeToken(TokenType.Then, ((IfBlockStatement)statement).ThenLocation);
            Serialize(((IfBlockStatement)statement).Statements);
            Serialize(((IfBlockStatement)statement).ElseIfBlockStatements);
            Serialize(((IfBlockStatement)statement).ElseBlockStatement);
            Serialize(((IfBlockStatement)statement).EndStatement);
        }

        private void SerializeElseBlockStatement(Tree statement)
        {
            SerializeComments(statement);
            foreach (Tree child in statement.Children)
                Serialize(child);
        }

        private void SerializeElseIfStatement(Tree statement)
        {
            SerializeComments(statement);
            Serialize(((ElseIfStatement)statement).Expression);
            SerializeToken(TokenType.Then, ((ElseIfStatement)statement).ThenLocation);
        }

        private void SerializeElseIfBlockStatement(Tree statement)
        {
            SerializeComments(statement);
            foreach (Tree child in statement.Children)
                Serialize(child);
        }

        private void SerializeLineIfBlockStatement(Tree statement)
        {
            SerializeComments(statement);
            Serialize(((LineIfStatement)statement).Expression);
            SerializeToken(TokenType.Then, ((LineIfStatement)statement).ThenLocation);
            Serialize(((LineIfStatement)statement).IfStatements);
            SerializeToken(TokenType.Else, ((LineIfStatement)statement).ElseLocation);
            Serialize(((LineIfStatement)statement).ElseStatements);
        }

        private void SerializeEmptyStatement(Tree tree)
        {
            SerializeComments(tree);
            foreach (Tree child in tree.Children)
                Serialize(child);
        }

        private void SerializeEndStatement(Tree statement)
        {
            SerializeComments(statement);
            foreach (Tree child in statement.Children)
                Serialize(child);
        }

        #endregion

        #region Declarations

        private void SerializeDeclaration(Tree declaration)
        {
            switch (declaration.Type)
            {
                case TreeType.EndBlockDeclaration:
                    SerializeEndBlockDeclaration(declaration);
                    break;

                case TreeType.EventDeclaration:
                    SerializeEventDeclaration(declaration);
                    break;

                case TreeType.CustomEventDeclaration:
                    SerializeCustomEventDeclaration(declaration);
                    break;

                case TreeType.ConstructorDeclaration:
                case TreeType.SubDeclaration:
                case TreeType.FunctionDeclaration:
                    SerializeMethodDeclaration(declaration);
                    break;

                case TreeType.OperatorDeclaration:
                    SerializeOperatorDeclaration(declaration);
                    break;

                case TreeType.ExternalSubDeclaration:
                case TreeType.ExternalFunctionDeclaration:
                    SerializeExternalDeclaration(declaration);
                    break;

                case TreeType.PropertyDeclaration:
                    SerializePropertyDeclaration(declaration);
                    break;

                case TreeType.GetAccessorDeclaration:
                    SerializeGetAccessorDeclaration(declaration);
                    break;

                case TreeType.SetAccessorDeclaration:
                    SerializeSetAccessorDeclaration(declaration);
                    break;

                case TreeType.EnumDeclaration:
                    SerializeEnumDeclaration(declaration);
                    break;

                case TreeType.EnumValueDeclaration:
                    SerializeEnumValueDeclaration(declaration);
                    break;

                case TreeType.DelegateSubDeclaration:
                    SerializeDelegateSubDeclaration(declaration);
                    break;

                case TreeType.DelegateFunctionDeclaration:
                    SerializeDelegateFunctionDeclaration(declaration);
                    break;

                case TreeType.ModuleDeclaration:
                    SerializeModuleDeclaration(declaration);
                    break;

                case TreeType.ClassDeclaration:
                    SerializeClassDeclaration(declaration);
                    break;
                
                case TreeType.StructureDeclaration:
                    SerializeStructureDeclaration(declaration);
                    break;

                case TreeType.InterfaceDeclaration:
                    SerializeInterfaceDeclaration(declaration);
                    break;

                case TreeType.NamespaceDeclaration:
                    SerializeNamespaceDeclaration(declaration);
                    break;

                case TreeType.OptionDeclaration:
                    SerializeOptionDeclaration(declaration);
                    break;

                case TreeType.VariableListDeclaration:
                    SerializeVariableListDeclaration(declaration);
                    break;

                case TreeType.EmptyDeclaration:
                    SerializeEmptyDeclaration(declaration);
                    break;

                case TreeType.ImplementsDeclaration:
                    SerializeImplementsDeclaration(declaration);
                    break;

                default:
                    SerializeComments(declaration);
                    foreach (Tree child in declaration.Children)
                        Serialize(child);
                    break;
            }
        }

        private void SerializeEndBlockDeclaration(Tree declaration)
        {
            Output.WriteAttributeString("endType", ((EndBlockDeclaration)declaration).EndType.ToString());
            SerializeComments(declaration);
            SerializeToken(GetBlockTypeToken(((EndBlockDeclaration)declaration).EndType), ((EndBlockDeclaration)declaration).EndArgumentLocation);
        }

        private void SerializeEventDeclaration(Tree declaration)
        {
            SerializeComments(declaration);
            Serialize(((EventDeclaration)declaration).Attributes);
            Serialize(((EventDeclaration)declaration).Modifiers);
            SerializeToken(TokenType.Event, ((EventDeclaration)declaration).KeywordLocation);
            Serialize(((EventDeclaration)declaration).Name);
            Serialize(((EventDeclaration)declaration).Parameters);
            SerializeToken(TokenType.As, ((EventDeclaration)declaration).AsLocation);
            Serialize(((EventDeclaration)declaration).ResultTypeAttributes);
            Serialize(((EventDeclaration)declaration).ResultType);
            Serialize(((EventDeclaration)declaration).ImplementsList);
        }

        private void SerializeCustomEventDeclaration(Tree declaration)
        {
            SerializeComments(declaration);
            Serialize(((CustomEventDeclaration)declaration).Attributes);
            Serialize(((CustomEventDeclaration)declaration).Modifiers);
            SerializeToken(TokenType.Custom, ((CustomEventDeclaration)declaration).CustomLocation);
            SerializeToken(TokenType.Event, ((CustomEventDeclaration)declaration).KeywordLocation);
            Serialize(((CustomEventDeclaration)declaration).Name);
            SerializeToken(TokenType.As, ((CustomEventDeclaration)declaration).AsLocation);
            Serialize(((CustomEventDeclaration)declaration).ResultType);
            Serialize(((CustomEventDeclaration)declaration).ImplementsList);
            Serialize(((CustomEventDeclaration)declaration).Accessors);
            Serialize(((CustomEventDeclaration)declaration).EndDeclaration);
        }

        private void SerializeMethodDeclaration(Tree declaration)
        {
            SerializeComments(declaration);
            Serialize(((MethodDeclaration)declaration).Attributes);
            Serialize(((MethodDeclaration)declaration).Modifiers);

            switch (declaration.Type)
            {
                case TreeType.SubDeclaration:
                    SerializeToken(TokenType.Sub, ((MethodDeclaration)declaration).KeywordLocation);
                    break;

                case TreeType.FunctionDeclaration:
                    SerializeToken(TokenType.Function, ((MethodDeclaration)declaration).KeywordLocation);
                    break;

                case TreeType.ConstructorDeclaration:
                    SerializeToken(TokenType.New, ((MethodDeclaration)declaration).KeywordLocation);
                    break;
            }

            Serialize(((MethodDeclaration)declaration).Name);
            Serialize(((MethodDeclaration)declaration).Parameters);
            Serialize(((MethodDeclaration)declaration).TypeParameters);
            SerializeToken(TokenType.As, ((MethodDeclaration)declaration).AsLocation);
            Serialize(((MethodDeclaration)declaration).ResultTypeAttributes);
            Serialize(((MethodDeclaration)declaration).ResultType);
            Serialize(((MethodDeclaration)declaration).ImplementsList);
            Serialize(((MethodDeclaration)declaration).HandlesList);
            Serialize(((MethodDeclaration)declaration).Statements);
            Serialize(((MethodDeclaration)declaration).EndDeclaration);
        }

        private void SerializeOperatorDeclaration(Tree declaration)
        {
            SerializeComments(declaration);
            Serialize(((OperatorDeclaration)declaration).Attributes);
            Serialize(((OperatorDeclaration)declaration).Modifiers);
            SerializeToken(TokenType.Operator, ((OperatorDeclaration)declaration).KeywordLocation);
            if (((OperatorDeclaration)declaration).OperatorToken != null)
                SerializeToken(((OperatorDeclaration)declaration).OperatorToken.Type, ((OperatorDeclaration)declaration).OperatorToken.Span.Start);

            Serialize(((OperatorDeclaration)declaration).Parameters);
            SerializeToken(TokenType.As, ((OperatorDeclaration)declaration).AsLocation);
            Serialize(((OperatorDeclaration)declaration).ResultTypeAttributes);
            Serialize(((OperatorDeclaration)declaration).ResultType);
            Serialize(((OperatorDeclaration)declaration).Statements);
            Serialize(((OperatorDeclaration)declaration).EndDeclaration);
        }

        private void SerializeExternalDeclaration(Tree declaration)
        {
            SerializeComments(declaration);
            Serialize(((ExternalDeclaration)declaration).Attributes);
            Serialize(((ExternalDeclaration)declaration).Modifiers);

            SerializeToken(TokenType.Declare, ((ExternalDeclaration)declaration).KeywordLocation);

            switch (((ExternalDeclaration)declaration).Charset)
            {
                case Charset.Auto:
                    SerializeToken(TokenType.Auto, ((ExternalDeclaration)declaration).CharsetLocation);
                    break;

                case Charset.Ansi:
                    SerializeToken(TokenType.Ansi, ((ExternalDeclaration)declaration).CharsetLocation);
                    break;

                case Charset.Unicode:
                    SerializeToken(TokenType.Unicode, ((ExternalDeclaration)declaration).CharsetLocation);
                    break;
            }

            switch (declaration.Type)
            {
                case TreeType.ExternalSubDeclaration:
                    SerializeToken(TokenType.Sub, ((ExternalDeclaration)declaration).SubOrFunctionLocation);
                    break;

                case TreeType.ExternalFunctionDeclaration:
                    SerializeToken(TokenType.Function, ((ExternalDeclaration)declaration).SubOrFunctionLocation);
                    break;
            }

            Serialize(((ExternalDeclaration)declaration).Name);
            SerializeToken(TokenType.Lib, ((ExternalDeclaration)declaration).LibLocation);
            Serialize(((ExternalDeclaration)declaration).LibLiteral);
            SerializeToken(TokenType.Alias, ((ExternalDeclaration)declaration).AliasLocation);
            Serialize(((ExternalDeclaration)declaration).AliasLiteral);
            Serialize(((ExternalDeclaration)declaration).Parameters);
            SerializeToken(TokenType.As, ((ExternalDeclaration)declaration).AsLocation);
            Serialize(((ExternalDeclaration)declaration).ResultTypeAttributes);
            Serialize(((ExternalDeclaration)declaration).ResultType);
        }

        private void SerializePropertyDeclaration(Tree declaration)
        {
            PropertyDeclaration Result = (PropertyDeclaration)declaration;
            SerializeComments(Result);
            Serialize(Result.Attributes);
            Serialize(Result.Modifiers);
            SerializeToken(TokenType.Property, Result.KeywordLocation);
            Serialize(Result.Name);
            Serialize(Result.Parameters);
            SerializeToken(TokenType.As, Result.AsLocation);
            Serialize(Result.ResultTypeAttributes);
            Serialize(Result.ResultType);
            Serialize(Result.ImplementsList);
            Serialize(Result.GetAccessor);
            Serialize(Result.SetAccessor);
            Serialize(Result.Statements);
            Serialize(Result.EndDeclaration);
        }

        private void SerializeGetAccessorDeclaration(Tree declaration)
        {
            SerializeComments(declaration);
            Serialize(((GetAccessorDeclaration)declaration).Attributes);
            Serialize(((GetAccessorDeclaration)declaration).Modifiers);
            SerializeToken(TokenType.Get, ((GetAccessorDeclaration)declaration).GetLocation);
            Serialize(((GetAccessorDeclaration)declaration).Statements);
            Serialize(((GetAccessorDeclaration)declaration).EndDeclaration);
        }

        private void SerializeSetAccessorDeclaration(Tree declaration)
        {
            SerializeComments(declaration);
            Serialize(((SetAccessorDeclaration)declaration).Attributes);
            Serialize(((SetAccessorDeclaration)declaration).Modifiers);
            SerializeToken(TokenType.Set, ((SetAccessorDeclaration)declaration).SetLocation);
            Serialize(((SetAccessorDeclaration)declaration).Parameters);
            Serialize(((SetAccessorDeclaration)declaration).Statements);
            Serialize(((SetAccessorDeclaration)declaration).EndDeclaration);
        }

        private void SerializeEnumDeclaration(Tree declaration)
        {
            SerializeComments(declaration);
            Serialize(((EnumDeclaration)declaration).Attributes);
            Serialize(((EnumDeclaration)declaration).Modifiers);
            SerializeToken(TokenType.Enum, ((EnumDeclaration)declaration).KeywordLocation);
            Serialize(((EnumDeclaration)declaration).Name);
            SerializeToken(TokenType.As, ((EnumDeclaration)declaration).AsLocation);
            Serialize(((EnumDeclaration)declaration).ElementType);
            Serialize(((EnumDeclaration)declaration).Declarations);
            Serialize(((EnumDeclaration)declaration).EndDeclaration);
        }

        private void SerializeEnumValueDeclaration(Tree declaration)
        {
            SerializeComments(declaration);
            Serialize(((EnumValueDeclaration)declaration).Attributes);
            Serialize(((EnumValueDeclaration)declaration).Modifiers);
            Serialize(((EnumValueDeclaration)declaration).Name);
            SerializeToken(TokenType.Equals, ((EnumValueDeclaration)declaration).EqualsLocation);
            Serialize(((EnumValueDeclaration)declaration).Expression);
        }

        private void SerializeDelegateSubDeclaration(Tree declaration)
        {
            SerializeComments(declaration);
            Serialize(((DelegateSubDeclaration)declaration).Attributes);
            Serialize(((DelegateSubDeclaration)declaration).Modifiers);
            SerializeToken(TokenType.Delegate, ((DelegateSubDeclaration)declaration).KeywordLocation);
            SerializeToken(TokenType.Sub, ((DelegateSubDeclaration)declaration).SubOrFunctionLocation);
            Serialize(((DelegateSubDeclaration)declaration).Name);
            Serialize(((DelegateSubDeclaration)declaration).TypeParameters);
            Serialize(((DelegateSubDeclaration)declaration).Parameters);
        }

        private void SerializeDelegateFunctionDeclaration(Tree declaration)
        {
            SerializeComments(declaration);
            Serialize(((DelegateFunctionDeclaration)declaration).Attributes);
            Serialize(((DelegateFunctionDeclaration)declaration).Modifiers);
            SerializeToken(TokenType.Delegate, ((DelegateFunctionDeclaration)declaration).KeywordLocation);
            SerializeToken(TokenType.Function, ((DelegateFunctionDeclaration)declaration).SubOrFunctionLocation);
            Serialize(((DelegateFunctionDeclaration)declaration).Name);
            Serialize(((DelegateFunctionDeclaration)declaration).Parameters);
            SerializeToken(TokenType.As, ((DelegateFunctionDeclaration)declaration).AsLocation);
            Serialize(((DelegateFunctionDeclaration)declaration).ResultTypeAttributes);
            Serialize(((DelegateFunctionDeclaration)declaration).ResultType);
        }

        private void SerializeModuleDeclaration(Tree declaration)
        {
            SerializeComments(declaration);
            Serialize(((ModuleDeclaration)declaration).Attributes);
            Serialize(((ModuleDeclaration)declaration).Modifiers);
            SerializeToken(TokenType.Module, ((ModuleDeclaration)declaration).KeywordLocation);
            Serialize(((ModuleDeclaration)declaration).Name);
            Serialize(((ModuleDeclaration)declaration).Declarations);
            Serialize(((ModuleDeclaration)declaration).EndDeclaration);
        }

        private void SerializeClassDeclaration(Tree declaration)
        {
            SerializeComments(declaration);
            Serialize(((ClassDeclaration)declaration).Attributes);
            Serialize(((ClassDeclaration)declaration).Modifiers);
            SerializeToken(TokenType.Class, ((ClassDeclaration)declaration).KeywordLocation);
            Serialize(((ClassDeclaration)declaration).Name);
            Serialize(((ClassDeclaration)declaration).TypeParameters);
            Serialize(((ClassDeclaration)declaration).Declarations);
            Serialize(((ClassDeclaration)declaration).EndDeclaration);
        }

        private void SerializeStructureDeclaration(Tree declaration)
        {
            SerializeComments(declaration);
            Serialize(((StructureDeclaration)declaration).Attributes);
            Serialize(((StructureDeclaration)declaration).Modifiers);
            SerializeToken(TokenType.Structure, ((StructureDeclaration)declaration).KeywordLocation);
            Serialize(((StructureDeclaration)declaration).Name);
            Serialize(((StructureDeclaration)declaration).TypeParameters);
            Serialize(((StructureDeclaration)declaration).Declarations);
            Serialize(((StructureDeclaration)declaration).EndDeclaration);
        }

        private void SerializeInterfaceDeclaration(Tree declaration)
        {
            SerializeComments(declaration);
            Serialize(((InterfaceDeclaration)declaration).Attributes);
            Serialize(((InterfaceDeclaration)declaration).Modifiers);
            SerializeToken(TokenType.Interface, ((InterfaceDeclaration)declaration).KeywordLocation);
            Serialize(((InterfaceDeclaration)declaration).Name);
            Serialize(((InterfaceDeclaration)declaration).TypeParameters);
            Serialize(((InterfaceDeclaration)declaration).Declarations);
            Serialize(((InterfaceDeclaration)declaration).EndDeclaration);
        }

        private void SerializeNamespaceDeclaration(Tree declaration)
        {
            SerializeComments(declaration);
            Serialize(((NamespaceDeclaration)declaration).Attributes);
            Serialize(((NamespaceDeclaration)declaration).Modifiers);

            SerializeToken(TokenType.Namespace, ((NamespaceDeclaration)declaration).NamespaceLocation);
            Serialize(((NamespaceDeclaration)declaration).Name);
            Serialize(((NamespaceDeclaration)declaration).Declarations);
            Serialize(((NamespaceDeclaration)declaration).EndDeclaration);
        }

        private void SerializeOptionDeclaration(Tree declaration)
        {
            Output.WriteAttributeString("type", ((OptionDeclaration)declaration).OptionType.ToString());
            SerializeComments(declaration);

            switch (((OptionDeclaration)declaration).OptionType)
            {
                case OptionType.SyntaxError:
                    break;

                case OptionType.Explicit:
                case OptionType.ExplicitOn:
                case OptionType.ExplicitOff:
                    Output.WriteStartElement("Explicit");
                    SerializePosition(((OptionDeclaration)declaration).OptionTypeLocation);
                    Output.WriteEndElement();

                    if (((OptionDeclaration)declaration).OptionType == OptionType.ExplicitOn)
                    {
                        SerializeToken(TokenType.On, ((OptionDeclaration)declaration).OptionArgumentLocation);
                    }
                    else if (((OptionDeclaration)declaration).OptionType == OptionType.ExplicitOff)
                    {
                        Output.WriteStartElement("Off");
                        SerializePosition(((OptionDeclaration)declaration).OptionArgumentLocation);
                        Output.WriteEndElement();
                    }

                    break;

                case OptionType.Strict:
                case OptionType.StrictOn:
                case OptionType.StrictOff:

                    Output.WriteStartElement("Strict");
                    SerializePosition(((OptionDeclaration)declaration).OptionTypeLocation);
                    Output.WriteEndElement();

                    if (((OptionDeclaration)declaration).OptionType == OptionType.StrictOn)
                    {
                        SerializeToken(TokenType.On, ((OptionDeclaration)declaration).OptionArgumentLocation);
                    }
                    else if (((OptionDeclaration)declaration).OptionType == OptionType.StrictOff)
                    {
                        Output.WriteStartElement("Off");
                        SerializePosition(((OptionDeclaration)declaration).OptionArgumentLocation);
                        Output.WriteEndElement();
                    }

                    break;

                case OptionType.CompareBinary:
                case OptionType.CompareText:
                    Output.WriteStartElement("Compare");
                    SerializePosition(((OptionDeclaration)declaration).OptionTypeLocation);
                    Output.WriteEndElement();

                    if (((OptionDeclaration)declaration).OptionType == OptionType.CompareBinary)
                    {
                        Output.WriteStartElement("Binary");
                        SerializePosition(((OptionDeclaration)declaration).OptionArgumentLocation);
                        Output.WriteEndElement();
                    }
                    else
                    {
                        Output.WriteStartElement("Text");
                        SerializePosition(((OptionDeclaration)declaration).OptionArgumentLocation);
                        Output.WriteEndElement();
                    }

                    break;
                case OptionType.BaseZero:
                case OptionType.BaseOne:
                    Output.WriteStartElement("Base");
                    SerializePosition(((OptionDeclaration)declaration).OptionTypeLocation);
                    Output.WriteEndElement();

                    if (((OptionDeclaration)declaration).OptionType == OptionType.BaseZero)
                    {
                        Output.WriteStartElement("Zero");
                        SerializePosition(((OptionDeclaration)declaration).OptionArgumentLocation);
                        Output.WriteEndElement();
                    }
                    else
                    {
                        Output.WriteStartElement("One");
                        SerializePosition(((OptionDeclaration)declaration).OptionArgumentLocation);
                        Output.WriteEndElement();
                    }

                    break;
            }
        }

        private void SerializeVariableListDeclaration(Tree declaration)
        {
            SerializeComments(declaration);
            foreach (Tree child in declaration.Children)
                Serialize(child);
        }

        private void SerializeEmptyDeclaration(Tree declaration)
        {
            SerializeComments(declaration);
            foreach (Tree child in declaration.Children)
                Serialize(child);
        }

        private void SerializeImplementsDeclaration(Tree declaration)
        {
            SerializeComments(declaration);
            foreach (Tree child in declaration.Children)
                Serialize(child);
        }

        #endregion

        #region Public Methods

        public override void Serialize(Tree tree)
        {
            if (tree == null)
                return;

            Output.WriteStartElement(tree.Type.ToString());

            if (tree.IsBad)
                Output.WriteAttributeString("isBad", true.ToString());

            SerializePosition(tree.Span);

            if (Tree.Types.IsCollection(tree))
            {
                SerializeCollection(tree);
            }
            else if (Tree.Types.IsComment(tree))
            {
                SerializeComment(tree);
            }
            else if (Tree.Types.IsName(tree))
            {
                SerializeName(tree);
            }
            else if (Tree.Types.IsType(tree))
            {
                SerializeTypeName(tree);
            }
            else if (Tree.Types.IsArgument(tree))
            {
                SerializeArgument(tree);
            }
            else if (Tree.Types.IsExpression(tree))
            {
                SerializeExpression(tree);
            }
            else if (Tree.Types.IsStatement(tree))
            {
                SerializeStatement(tree);
            }
            else if (Tree.Types.IsModifier(tree))
            {
                SerializeModifier(tree);
            }
            else if (Tree.Types.IsVariableDeclarator(tree))
            {
                SerializeVariableDeclarator(tree);
            }
            else if (Tree.Types.IsCaseClause(tree))
            {
                SerializeCaseClause(tree);
            }
            else if (Tree.Types.IsAttribute(tree))
            {
                SerializeAttribute(tree);
            }
            else if (Tree.Types.IsDeclaration(tree))
            {
                SerializeDeclaration(tree);
            }
            else if (Tree.Types.IsParameter(tree))
            {
                SerializeParameter(tree);
            }
            else if (Tree.Types.IsTypeParameter(tree))
            {
                SerializeTypeParameter(tree);
            }
            else if (Tree.Types.IsImport(tree))
            {
                SerializeImport(tree);
            }
            else
            {
                foreach (Tree child in tree.Children)
                    Serialize(child);
            }

            Output.WriteEndElement();
        }

        #endregion
    }
}

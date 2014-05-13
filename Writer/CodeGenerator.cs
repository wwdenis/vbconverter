using System;
using System.Collections.Generic;

using System.Text;
using System.IO;

using System.Collections;
using System.Runtime.InteropServices;

using System.CodeDom;
using System.CodeDom.Compiler;

using Microsoft.CSharp;
using Microsoft.VisualBasic;

using VBConverter.CodeWriter.Entities;
using VBConverter.CodeWriter.Transform;

using VBConverter.CodeParser;
using VBConverter.CodeParser.Base;
using VBConverter.CodeParser.Error;
using VBConverter.CodeParser.Serializers;
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

namespace VBConverter.CodeWriter
{
    public sealed class CodeGenerator
    {
        #region Fields
        
        private LanguageVersion _version;
        private List<OptionType> _visualBasicOptions = null;


        #endregion

        #region Properties

        public LanguageVersion Version
        {
            get 
            { 
                return _version; 
            }
            private set 
            {
                if (!Enum.IsDefined(typeof(LanguageVersion), value))
                    throw new ArgumentNullException("Version");

                _version = value; 
            }
        }

        public List<OptionType> VisualBasicOptions
        {
            get { return _visualBasicOptions; }
            private set { _visualBasicOptions = value; }
        }

        #endregion

        #region Constructor

        public CodeGenerator(LanguageVersion version)
        {
            this.Version = version;
            this.VisualBasicOptions = new List<OptionType>();
        }

        public CodeGenerator()
        {
        }

        #endregion

        #region Private Methods

        private bool HasOptionType(OptionType option)
        {
            foreach (OptionType item in VisualBasicOptions)
                if (item == option)
                    return true;

            return false;
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
                    throw new NotImplementedException("GetCompoundAssignmentOperatorToken not implemented the BinaryOperatorType " + compoundOperator.ToString());
            }

            // return TokenType.None;
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
                    throw new NotImplementedException("GetBlockTypeToken not implemented the BlockType " + blockType.ToString());
            }

            // return TokenType.None;
        }

        private static CodeStatementCollection ConvertComments(CodeCommentStatementCollection comments)
        {
            if (comments == null)
                return null;

            CodeStatementCollection result = new CodeStatementCollection();
            foreach (CodeCommentStatement item in comments)
                result.Add(item);

            return result;
        }

        private static CodeExpression[] ConvertCodeExpressionCollection(CodeExpressionCollection expressions)
        {
            if (expressions == null)
                return null;

            List<CodeExpression> arguments = new List<CodeExpression>();
            foreach (CodeExpression item in expressions)
                arguments.Add(item);

            return arguments.ToArray();
        }

        #endregion

        #region Helpers

        private MemberAttributes TranslateModifier(ModifierTypes modifier)
        {
            MemberAttributes output = MemberAttributes.Public;

            switch (modifier)
            {
                case ModifierTypes.Optional:
                    output = (MemberAttributes)0;
                    break;
                case ModifierTypes.ParamArray:
                    output = (MemberAttributes)0;
                    break;
                case ModifierTypes.ByVal:
                    output = (MemberAttributes)0;
                    break;
                case ModifierTypes.ByRef:
                    output = (MemberAttributes)0;
                    break;
                case ModifierTypes.WithEvents:
                    output = (MemberAttributes)0;
                    break;
                
                case ModifierTypes.Const:
                    output = MemberAttributes.Const;
                    break;

                case ModifierTypes.Dim:
                    output = MemberAttributes.Private;
                    break;
                case ModifierTypes.Public:
                    output = MemberAttributes.Public;
                    break;
                case ModifierTypes.Private:
                    output = MemberAttributes.Private;
                    break;
                case ModifierTypes.Protected:
                    output = MemberAttributes.Family;
                    break;
                case ModifierTypes.Friend:
                    output = MemberAttributes.FamilyAndAssembly;
                    break;
                
                case ModifierTypes.Static:
                    break;
                case ModifierTypes.Shared:
                    break;
                case ModifierTypes.Shadows:
                    break;
                case ModifierTypes.Overloads:
                    break;
                case ModifierTypes.MustInherit:
                    break;
                case ModifierTypes.NotInheritable:
                    break;
                case ModifierTypes.Overrides:
                    break;
                case ModifierTypes.NotOverridable:
                    break;
                case ModifierTypes.Overridable:
                    break;
                case ModifierTypes.MustOverride:
                    break;
                case ModifierTypes.ReadOnly:
                    break;
                case ModifierTypes.WriteOnly:
                    break;
                case ModifierTypes.Default:
                    break;
                case ModifierTypes.Partial:
                    break;
                case ModifierTypes.Widening:
                    break;
                case ModifierTypes.Narrowing:
                    break;
            }

            return output;
        }

        private Type TranslateIntrinsicType(IntrinsicType type)
        {
            Type output = typeof(object);

            switch (type)
            {
                case IntrinsicType.None:
                    output = typeof(object);
                    break;
                case IntrinsicType.Boolean:
                    output = typeof(bool);
                    break;
                case IntrinsicType.SByte:
                    output = typeof(sbyte);
                    break;
                case IntrinsicType.Byte:
                    output = typeof(byte);
                    break;
                case IntrinsicType.Short:
                    output = typeof(short);
                    break;
                case IntrinsicType.UShort:
                    output = typeof(ushort);
                    break;
                case IntrinsicType.Integer:
                    if (LanguageInfo.ImplementsVB60(Version))
                        output = typeof(short);
                    else
                        output = typeof(int);
                    break;
                case IntrinsicType.UInteger:
                    output = typeof(uint);
                    break;
                case IntrinsicType.Long:
                    if (LanguageInfo.ImplementsVB60(Version))
                        output = typeof(int);
                    else
                        output = typeof(long);
                    break;
                case IntrinsicType.ULong:
                    output = typeof(ulong);
                    break;
                case IntrinsicType.Decimal:
                    output = typeof(decimal);
                    break;
                case IntrinsicType.Single:
                    output = typeof(Single);
                    break;
                case IntrinsicType.Double:
                    output = typeof(double);
                    break;
                case IntrinsicType.Date:
                    output = typeof(DateTime);
                    break;
                case IntrinsicType.Char:
                    output = typeof(char);
                    break;
                case IntrinsicType.String:
                    output = typeof(string);
                    break;
                case IntrinsicType.Object:
                    output = typeof(object);
                    break;
                case IntrinsicType.Variant:
                    output = typeof(object);
                    break;
                case IntrinsicType.Currency:
                    output = typeof(decimal);
                    break;
                case IntrinsicType.FixedString:
                    output = typeof(string);
                    break;
            }

            return output;
        }

        private Type TranslateTypeCharacter(TypeCharacter type)
        {
            Type output = typeof(object);

            switch (type)
            {
                case TypeCharacter.StringSymbol:
                    output = typeof(string);
                    break;
                case TypeCharacter.IntegerSymbol:
                    if (LanguageInfo.ImplementsVB60(Version))
                        output = typeof(short);
                    else
                        output = typeof(int);
                    break;
                case TypeCharacter.LongSymbol:
                    if (LanguageInfo.ImplementsVB60(Version))
                        output = typeof(int);
                    else
                        output = typeof(long);
                    break;
                case TypeCharacter.ShortChar:
                    output = typeof(short);
                    break;
                case TypeCharacter.IntegerChar:
                    output = typeof(int);
                    break;
                case TypeCharacter.LongChar:
                    output = typeof(int);
                    break;
                case TypeCharacter.SingleSymbol:
                    output = typeof(Single);
                    break;
                case TypeCharacter.DoubleSymbol:
                    output = typeof(double);
                    break;
                case TypeCharacter.DecimalSymbol:
                    output = typeof(decimal);
                    break;
                case TypeCharacter.SingleChar:
                    output = typeof(Single);
                    break;
                case TypeCharacter.DoubleChar:
                    output = typeof(double);
                    break;
                case TypeCharacter.DecimalChar:
                    output = typeof(decimal);
                    break;
                case TypeCharacter.UnsignedShortChar:
                    output = typeof(ushort);
                    break;
                case TypeCharacter.UnsignedIntegerChar:
                    output = typeof(uint);
                    break;
                case TypeCharacter.UnsignedLongChar:
                    output = typeof(ulong);
                    break;
                default:
                    throw new NotImplementedException("TranslateTypeCharacter not implemented for the type character " + type.ToString());
            }

            return output;
        }

        private CodeExpression CreateArrayInitializer(CodeTypeReference arrayType, CodeExpressionCollection lengths)
        {
            //(int[,])Array.CreateInstance(typeof(int), 5, 10);
            CodeExpression output = null;

            if (lengths == null || lengths.Count <= 1)
            {
                CodeArrayCreateExpression create = new CodeArrayCreateExpression(arrayType);
                CodeExpression result = null;

                if (lengths != null && lengths.Count > 0)
                {
                    CodeExpression expression = lengths[0];
                    CodePrimitiveExpression length = expression as CodePrimitiveExpression;
                    
                    if (length != null)
                    {
                        long value = 0;
                        if (long.TryParse(length.Value.ToString(), out value))
                            result = new CodePrimitiveExpression(value + 1);
                    }
                    
                    if (result == null)
                        result = new CodeBinaryOperatorExpression(expression, CodeBinaryOperatorType.Add, new CodePrimitiveExpression(1));

                    create.SizeExpression = result;
                }

                output = create;
            }
            else
            {
                CodeVariableReferenceExpression variable = new CodeVariableReferenceExpression("Array");
                CodeMethodReferenceExpression method = new CodeMethodReferenceExpression(variable, "CreateInstance");

                CodeTypeReference elementType = new CodeTypeReference(arrayType, 0);
                CodeTypeOfExpression typeOf = new CodeTypeOfExpression(elementType);

                CodeExpressionCollection arguments = new CodeExpressionCollection();
                arguments.Add(typeOf);
                arguments.AddRange(lengths);
                CodeExpression[] argumentArray = ConvertCodeExpressionCollection(arguments);

                CodeMethodInvokeExpression invoke = new CodeMethodInvokeExpression(method, argumentArray);
                
                CodeCastExpression cast = new CodeCastExpression(arrayType, invoke);

                output = cast;
            }

            return output;
        }

        private void AddReturnStatements(CodeTypeReference returnType, string name, CodeStatementCollection statements)
        {
            string returnName = "_result_" + name;
            CodeExpression init = new CodeDefaultValueExpression(returnType);
            CodeVariableReferenceExpression returnVariable = new CodeVariableReferenceExpression(returnName);
            CodeVariableDeclarationStatement returnDeclaration = new CodeVariableDeclarationStatement(returnType, returnName, init);
            CodeMethodReturnStatement returnStatement = new CodeMethodReturnStatement(returnVariable);

            foreach (CodeStatement item in statements)
            {
                if (item is CodeAssignStatement)
                {
                    CodeAssignStatement assign = (CodeAssignStatement)item;
                    if (assign.Left is CodeVariableReferenceExpression)
                    {
                        CodeVariableReferenceExpression variable = (CodeVariableReferenceExpression)assign.Left;
                        if (string.Compare(variable.VariableName, name, true) == 0)
                            variable.VariableName = returnName;
                    }
                }
            }

            if (statements != null && statements.Count > 0)
            {
                statements.Insert(0, returnDeclaration);
                statements.Add(returnStatement);
            }
        }

        private CodeExpression GetIntrinsicInitializer(TypeName typeName)
        {
            CodeExpression initializer = null;
            if (typeName is IntrinsicTypeName)
            {
                IntrinsicTypeName type = (IntrinsicTypeName)typeName;
                if (type.IntrinsicType == IntrinsicType.FixedString)
                {
                    CodeExpression length = SerializeExpression(type.StringLength);
                    initializer = new CodeObjectCreateExpression("FixedLengthString", length);
                }
            }
            return initializer;
        }

        private CodeExpression GetExplicitArraySize(BinaryOperatorExpression input)
        {
            if (!(input.Parent != null && input.Parent is Argument &&
                input.Parent.Parent != null && input.Parent.Parent is ArgumentCollection &&
                input.Parent.Parent.Parent != null && input.Parent.Parent.Parent is ArrayTypeName &&
                input.Operator == BinaryOperatorType.To
                ))
                return null;

            CodeExpression left = SerializeExpression(input.LeftOperand);
            CodeExpression right = SerializeExpression(input.RightOperand);

            CodeExpression output = null;

            if (left is CodePrimitiveExpression && right is CodePrimitiveExpression)
            {
                int leftValue = 0;
                int rightValue = 0;
                if (int.TryParse(((CodePrimitiveExpression)left).Value.ToString(), out leftValue) &&
                    int.TryParse(((CodePrimitiveExpression)right).Value.ToString(), out rightValue)
                    )
                {
                    int length = rightValue - leftValue;
                    output = new CodePrimitiveExpression(length);
                }
            }

            if (output == null)
                output = new CodeBinaryOperatorExpression(right, CodeBinaryOperatorType.Subtract, left);
            
            return output;
        }

        #endregion

        #region Individual Trees

        private CodeExpression SerializeArgument(Tree tree)
        {
            if (tree == null)
                return new CodePrimitiveExpression(null);

            Argument input = (Argument)tree;
            CodeExpression output = SerializeExpression(input.Expression);
            return output;
        }

        private MemberAttributes SerializeModifier(Tree tree)
        {
            if (tree == null)
                return MemberAttributes.Public;

            Modifier input = (Modifier)tree;
            MemberAttributes output = TranslateModifier(input.ModifierType);

            return output;
        }

        private CodeVariableDeclaratorCollection SerializeVariableDeclarator(Tree tree)
        {
            if (tree == null)
                return new CodeVariableDeclaratorCollection(CodeCollectionType.TypeMember);
            
            VariableDeclarator input = (VariableDeclarator)tree;
            CodeVariableDeclaratorCollection output = null;

            if (input.IsField)
                output = new CodeVariableDeclaratorCollection(CodeCollectionType.TypeMember);
            else
                output = new CodeVariableDeclaratorCollection(CodeCollectionType.Statement);

            //input.Arguments

            CodeReferenceNameCollection variables = SerializeVariableNameCollection(input.VariableNames);
            CodeTypeReference type = SerializeTypeName(input.VariableType);
            CodeExpression initializer = SerializeInitializer(input.Initializer);

            CodeExpression intrinsicInitializer = GetIntrinsicInitializer(input.VariableType);
            if (intrinsicInitializer != null)
                initializer = intrinsicInitializer;


            foreach (CodeReferenceName variable in variables)
            {
                CodeExpression arrayInitializer = null;
                CodeTypeReference variableType = type;

                if (variable.Rank > 0)
                {
                    variableType = new CodeTypeReference(type, variable.Rank);
                    arrayInitializer = CreateArrayInitializer(variableType, variable.ArrayLengths);

                    if (arrayInitializer != null)
                        initializer = arrayInitializer;
                }

                if (input.IsField)
                {    
                    CodeMemberField field = new CodeMemberField();
                    field.Name = variable.Name;
                    field.Type = variableType;
                    field.InitExpression = initializer;

                    output.Members.Add(field);
	            }
                else
                {
                    CodeVariableDeclarationStatement statement = new CodeVariableDeclarationStatement();
                    statement.Name = variable.Name;
                    statement.Type = variableType;
                    statement.InitExpression = initializer;

                    output.Statements.Add(statement);
                }
            }

            return output;
        }

        private void SerializeTypeParameter(Tree tree)
        {
            throw new NotImplementedException();
            //TypeParameter input = (TypeParameter)tree;
            //input.TypeName
            //input.TypeConstraints
        }

        private CodeParameter SerializeParameter(Tree tree)
        {
            if (tree == null)
                return null;

            Parameter input = (Parameter)tree;
            CodeParameter output = new CodeParameter();

            CodeParameterDeclarationExpression parameter = new CodeParameterDeclarationExpression();

            MemberAttributes modifiers = SerializeModifierCollection(input.Modifiers);
            CodeReferenceName name = SerializeName(input.VariableName);
            CodeTypeReference type = SerializeTypeName(input.ParameterType);
            CodeExpression initializer = SerializeInitializer(input.Initializer);

            // Assuming only the modifiers ByVal, ByRef, Optional and ParamArray
            bool isByRef = false;
            bool isByVal = false;
            bool isOptional = false;
            bool isParamArray = false;

            if (input.Modifiers != null)
            {
                isByRef = input.Modifiers.IsOfType(ModifierTypes.ByRef);
                isByVal = input.Modifiers.IsOfType(ModifierTypes.ByVal);
                isOptional = input.Modifiers.IsOfType(ModifierTypes.Optional);
                isParamArray = input.Modifiers.IsOfType(ModifierTypes.ParamArray);
            }

            // Assuming ByRef when there is no modifier
            if (LanguageInfo.ImplementsVB60(Version) && !isByVal)
                isByRef = true;

            if (isByRef)
                parameter.Direction = FieldDirection.Ref;

            if (isParamArray)
                parameter.CustomAttributes.Add(new CodeAttributeDeclaration(typeof(ParamArrayAttribute).FullName));

            parameter.Name = name.Name;
            parameter.Type = type;

            output.Parameter = parameter;
            output.IsOptional = isOptional;
            output.InitValue = initializer;

            //input.Attributes

            return output;
        }

        private void SerializeAttribute(Tree tree)
        {
            throw new NotImplementedException();
            //AttributeTree input = (AttributeTree)tree;
            //input.AttributeType
            //input.Name
            //input.Arguments
        }

        #endregion

        #region Comments

        private CodeCommentStatementCollection SerializeComments(IList<Comment> comments)
        {
            if (comments == null)
                return new CodeCommentStatementCollection();
            
            //CommentableTree input = (CommentableTree)tree;
            CodeCommentStatementCollection output = new CodeCommentStatementCollection();

            foreach (Comment item in comments)
            {
                CodeCommentStatement comment = SerializeComment(item);
                output.Add(comment);
            }

            return output;
        }

        private CodeCommentStatement SerializeComment(Tree tree)
        {
            if (tree == null)
                return null;

            string content = ((Comment)tree).Text;

            if (Version == LanguageVersion.VB6 && content.EndsWith("_"))
                content = content.Substring(0, content.Length - 1);

            CodeCommentStatement comment = new CodeCommentStatement(content);
            
            return comment;
        }

        #endregion

        #region Collections

        private MemberAttributes SerializeModifierCollection(Tree tree)
        {
            if (tree == null)
                return MemberAttributes.Public;

            ModifierCollection input = (ModifierCollection)tree;
            MemberAttributes output = (MemberAttributes)0;
            
            foreach (Modifier modifier in input.Children)
            {
                MemberAttributes item = SerializeModifier(modifier);
                output = output | item;
            }

            return output;
        }

        private CodeExpressionCollection SerializeArgumentCollection(Tree tree)
        {
            if (tree == null)
                return new CodeExpressionCollection();

            ArgumentCollection input = (ArgumentCollection)tree;
            CodeExpressionCollection output = new CodeExpressionCollection();
            CodeExpression expression = null;
            
            foreach (Argument item in input)
            {
                expression = SerializeArgument(item);
                output.Add(expression);
            }

            return output;
        }

        private void SerializeAttributeCollection(Tree tree)
        {
            throw new NotImplementedException();
            //AttributeCollection input = (AttributeCollection)tree;
            //foreach (AttributeTree item in input)
            //    SerializeAttribute(item);
        }

        private CodeExpressionCollection SerializeCaseClauseCollection(Tree tree)
        {
            if (tree == null)
                return new CodeExpressionCollection();

            CaseClauseCollection input = (CaseClauseCollection)tree;
            CodeExpressionCollection output = new CodeExpressionCollection();

            foreach (CaseClause item in input)
            {
                CodeExpression condition = SerializeCaseClause(item);
                output.Add(condition);
            }

            return output;
        }

        private CodeExpressionCollection SerializeExpressionCollection(Tree tree)
        {
            if (tree == null)
                return new CodeExpressionCollection();

            ExpressionCollection input = (ExpressionCollection)tree;
            CodeExpressionCollection output = new CodeExpressionCollection();

            foreach (Expression item in input)
            {
                CodeExpression expression = SerializeExpression(item);
                output.Add(expression);
            }

            return output;
        }

        private void SerializeImportCollection(Tree tree)
        {
            throw new NotImplementedException();
            //ImportCollection input = (ImportCollection)tree;
            //foreach (Import item in input)
            //    SerializeImport(item);
        }

        private CodeExpressionCollection SerializeInitializerCollection(Tree tree)
        {
            if (tree == null)
                return new CodeExpressionCollection();

            InitializerCollection input = (InitializerCollection)tree;
            CodeExpressionCollection output = new CodeExpressionCollection();

            foreach (Initializer item in input)
            {
                Expression initExpression = null;

                if (item is AggregateInitializer)
                    throw new NotImplementedException();
                else
                    initExpression = ((ExpressionInitializer)item).Expression;

                CodeExpression expression = SerializeExpression(initExpression);
                output.Add(expression);
            }

            return output;
        }

        private CodeReferenceNameCollection SerializeNameCollection(Tree tree)
        {
            if (tree == null)
                return null;

            NameCollection input = (NameCollection)tree;
            CodeReferenceNameCollection output = new CodeReferenceNameCollection();

            foreach (Name item in input)
            {
                CodeReferenceName name = SerializeName(item);
                output.Add(name);
            }

            return output;
        }

        private CodeReferenceNameCollection SerializeVariableNameCollection(Tree tree)
        {
            if (tree == null)
                return null;

            VariableNameCollection input = (VariableNameCollection)tree;
            CodeReferenceNameCollection output = new CodeReferenceNameCollection();

            foreach (VariableName item in input)
            {
                CodeReferenceName name = SerializeVariableName(item);
                output.Add(name);
            }

            return output;
        }

        private CodeParameterCollection SerializeParameterCollection(Tree tree)
        {
            if (tree == null)
                return new CodeParameterCollection();

            ParameterCollection input = (ParameterCollection)tree;
            CodeParameterCollection output = new CodeParameterCollection();

            foreach (Parameter item in input)
            {
                CodeParameter parameter = SerializeParameter(item);
                output.Add(parameter);
            }

            return output;
        }

        private CodeTypeReferenceCollection SerializeTypeNameCollection(Tree tree)
        {
            if (tree == null)
                return new CodeTypeReferenceCollection();

            TypeNameCollection input = (TypeNameCollection)tree;
            CodeTypeReferenceCollection output = new CodeTypeReferenceCollection();

            foreach (TypeName item in input)
            {
                CodeTypeReference type = SerializeTypeName(item);
                output.Add(type);
            }

            return output;
        }

        private CodeVariableDeclaratorCollection SerializeVariableDeclaratorCollection(Tree tree)
        {
            if (tree == null)
                return new CodeVariableDeclaratorCollection(CodeCollectionType.TypeMember);

            VariableDeclaratorCollection input = (VariableDeclaratorCollection)tree;
            CodeVariableDeclaratorCollection output = null;

            foreach (VariableDeclarator child in input.Children)
            {
                CodeVariableDeclaratorCollection item = SerializeVariableDeclarator(child);

                switch (item.CollectionType)
                {
                    case CodeCollectionType.TypeMember:
                        if (item.Statements.Count > 0)
                            throw new Exception("Invalid VariableDeclaratorCollection. This class have member fields and local variables declarations.");
                        output = new CodeVariableDeclaratorCollection(item.Members);                        
                        break;

                    case CodeCollectionType.Statement:
                        if (item.Members.Count > 0)
                            throw new Exception("Invalid VariableDeclaratorCollection. This class have member fields and local variables declarations.");
                        output = new CodeVariableDeclaratorCollection(item.Statements);
                        break;
                }
            }

            return output;
        }

        private CodeDeclarationCollection SerializeDeclarationCollection(Tree tree)
        {
            if (tree == null)
                return new CodeDeclarationCollection();

            DeclarationCollection input = (DeclarationCollection)tree;
            CodeDeclarationCollection output = new CodeDeclarationCollection();

            foreach (Declaration item in input)
            {
                CodeDeclaration declaration = SerializeDeclaration(item);
                output.Add(declaration);
            }

            return output;
        }

        private CodeStatementCollection SerializeStatementCollection(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            StatementCollection input = (StatementCollection)tree;
            CodeStatementCollection output = new CodeStatementCollection();

            foreach (Statement item in input)
            {
                CodeStatementCollection statement = SerializeStatement(item); ;
                output.AddRange(statement);
            }

            return output;
        }

        private CodeTypeParameterCollection SerializeTypeParameterCollection(Tree tree)
        {
            throw new NotImplementedException();
            //TypeParameterCollection input = (TypeParameterCollection)tree;
            //foreach (TypeParameter item in input)
            //    SerializeTypeParameter(item);
        }

        private CodeTypeParameterCollection SerializeTypeConstraintCollection(Tree tree)
        {
            throw new NotImplementedException();
            //TypeConstraintCollection input = (TypeConstraintCollection)tree;
            //foreach (TypeName item in input)
            //    SerializeTypeName(item);
        }

        private CodeTypeParameterCollection SerializeTypeArgumentCollection(Tree tree)
        {
            throw new NotImplementedException();
            //TypeArgumentCollection input = (TypeArgumentCollection)tree;
            //foreach (TypeName item in input)
            //    SerializeTypeName(item);
        }

        #endregion

        #region Names

        private CodeReferenceName SerializeName(Tree tree)
        {
            if (tree == null)
                return null;
            
            CodeReferenceName output = null;

            switch (tree.Type)
            {
                case TreeType.SimpleName:
                    output = SerializeSimpleName(tree);
                    break;

                case TreeType.QualifiedName:
                    output = SerializeQualifiedName(tree);
                    break;

                case TreeType.VariableName:
                    output = SerializeVariableName(tree);
                    break;
            }

            return output;
        }

        private CodeReferenceName SerializeSimpleName(Tree tree)
        {
            if (tree == null)
                return null;

            SimpleName input = (SimpleName)tree;
            CodeReferenceName output = new CodeReferenceName(input.Name);
            // TODO: Translate TypeCharacter to intrisic type
            //Type type = TranslateTypeCharacter(input.TypeCharacter);
            return output;
        }

        private CodeReferenceName SerializeQualifiedName(Tree tree)
        {
            if (tree == null)
                return null;

            QualifiedName input = (QualifiedName)tree;
            CodeReferenceName output = new CodeReferenceName(input.FullName);
            
            return output;
        }

        private CodeReferenceName SerializeVariableName(Tree tree)
        {
            if (tree == null)
                return null;

            VariableName input = (VariableName)tree;
            CodeReferenceName output = new CodeReferenceName();

            int rank = 0;
            CodeReferenceName name = SerializeSimpleName(input.Name);
            CodeExpressionCollection arrayLengths = null;

            if (input.ArrayType != null)
            {
                rank = input.ArrayType.Rank;
                arrayLengths = SerializeArgumentCollection(input.ArrayType.Arguments);
            }

            output.Name = name.Name;
            output.Rank = rank;
            output.ArrayLengths = arrayLengths;
            output.IsVariable = true;

            return output;
        }

        #endregion

        #region Types

        private CodeTypeReference SerializeTypeName(Tree tree)
        {
            if (tree == null)
                return new CodeTypeReference(typeof(object));

            CodeTypeReference type = null;

            switch (tree.Type)
            {
                case TreeType.IntrinsicType:
                    type = SerializeIntrinsicTypeName(tree);
                    break;

                case TreeType.NamedType:
                    type = SerializeNamedTypeName(tree);
                    break;
                case TreeType.ArrayType:
                    type = SerializeArrayTypeName(tree);
                    break;
                
                case TreeType.ConstructedType:
                    type = SerializeConstructedTypeName(tree);
                    break;
                
                default:
                    throw new NotImplementedException("SerializeType not implemented for class " + tree.GetType().Name);
            }

            return type;
        }

        private CodeTypeReference SerializeIntrinsicTypeName(Tree tree)
        {
            if (tree == null)
                return null;
            
            IntrinsicTypeName input = (IntrinsicTypeName)tree;
            //input.StringLength

            Type type = TranslateIntrinsicType(input.IntrinsicType);
            CodeTypeReference output = null;
            if (input.IntrinsicType == IntrinsicType.FixedString)
                output = new CodeTypeReference("FixedLengthString");
            else
                output = new CodeTypeReference(type);
            
            return output;
        }

        private CodeTypeReference SerializeNamedTypeName(Tree tree)
        {
            if (tree == null)
                return null;

            NamedTypeName input = (NamedTypeName)tree;
            CodeReferenceName name = SerializeName(input.Name);
            CodeTypeReference output = new CodeTypeReference(name.Name);

            return output;
        }

        private CodeTypeReference SerializeArrayTypeName(Tree tree)
        {
            if (tree == null)
                return null;

            ArrayTypeName input = (ArrayTypeName)tree;
            CodeTypeReference output = new CodeTypeReference(); ;
            
            CodeExpressionCollection arguments = SerializeArgumentCollection(input.Arguments);
            CodeTypeReference elementType = SerializeTypeName(input.ElementTypeName);

            output.ArrayElementType = elementType;
            output.ArrayRank = input.Rank;

            return output;
        }

        private CodeTypeReference SerializeConstructedTypeName(Tree tree)
        {
            throw new NotImplementedException();
            //Serialize(((ConstructedTypeName)Type).Name);
            //Serialize(((ConstructedTypeName)Type).TypeArguments);
        }

        #endregion

        #region Initializers

        private CodeExpression SerializeInitializer(Tree tree)
        {
            if (tree == null)
                return null;

            CodeExpression output = null;

            switch (tree.Type)
            {
                case TreeType.AggregateInitializer:
                    output = SerializeAggregateInitializer(tree);
                    break;

                case TreeType.ExpressionInitializer:
                    output = SerializeExpressionInitializer(tree);
                    break;

                default:
                    throw new NotImplementedException("SerializeInitializer not implemented for class " + tree.GetType().Name);
            }

            return output;
        }

        private CodeExpression SerializeAggregateInitializer(Tree tree)
        {
            throw new NotImplementedException();            
            //AggregateInitializer input = (AggregateInitializer)initializer;
            //input.Elements
        }

        private CodeExpression SerializeExpressionInitializer(Tree tree)
        {
            if (tree == null)
                return null;

            ExpressionInitializer input = (ExpressionInitializer)tree;
            CodeExpression output = SerializeExpression(input.Expression);

            return output;
        }

        #endregion

        #region Imports

        private CodeNamespaceImport SerializeImport(Tree Import)
        {
            throw new NotImplementedException();
            //switch (Import.Type)
            //{
            //    case TreeType.NameImport:
            //        SerializeNameImport(Import);
            //        break;

            //    case TreeType.AliasImport:
            //        SerializeAliasImport(Import);
            //        break;
            //}
        }

        private CodeNamespaceImport SerializeAliasImport(Tree import)
        {
            throw new NotImplementedException();
            //AliasImport input = (AliasImport)import;
            //input.AliasedTypeName
            //input.Name
        }

        private CodeNamespaceImport SerializeNameImport(Tree import)
        {
            throw new NotImplementedException();
            //NameImport input = (NameImport)import;
            //input.TypeName
        }

        #endregion

        #region Case Clause

        private CodeExpression SerializeCaseClause(Tree tree)
        {
            if (tree == null)
                return null;

            CodeExpression output = null;

            switch (tree.Type)
            {
                case TreeType.ComparisonCaseClause:
                    output = SerializeComparisonCaseClause(tree);
                    break;

                case TreeType.RangeCaseClause:
                    output = SerializeRangeCaseClause(tree);
                    break;

                default:
                    throw new NotImplementedException("SerializeCaseClause not implemented for class " + tree.GetType().Name);
            }

            return output;
        }

        private CodeExpression SerializeComparisonCaseClause(Tree tree)
        {
            if (tree == null)
                return null;

            ComparisonCaseClause input = (ComparisonCaseClause)tree;
            CodeExpression output = null;

            //Expression expression = input.Parent.Parent;
            //input.Operand;
            //input.ComparisonOperator

            output = SerializeBinaryOperatorExpression(input.ComparisonExpression);
            
            return output;
        }

        private CodeExpression SerializeRangeCaseClause(Tree tree)
        {
            if (tree == null)
                return null;

            RangeCaseClause input = (RangeCaseClause)tree;
            CodeExpression output = null;

            output = SerializeBinaryOperatorExpression(input.ComparisonExpression);

            return output;
        }

        #endregion

        #region Expressions

        private CodeExpression SerializeExpression(Tree tree)
        {
            if (tree == null)
                return new CodePrimitiveExpression(null);

            CodeExpression output = null;

            switch (tree.Type)
            {
                case TreeType.CallOrIndexExpression:
                    output = SerializeCallOrIndexExpression(tree);
                    break;

                case TreeType.BooleanLiteralExpression:
                case TreeType.StringLiteralExpression:
                case TreeType.CharacterLiteralExpression:
                case TreeType.DateLiteralExpression:
                case TreeType.IntegerLiteralExpression:
                case TreeType.FloatingPointLiteralExpression:
                case TreeType.DecimalLiteralExpression:
                    output = SerializeLiteralExpression(tree);
                    break;

                case TreeType.GetTypeExpression:
                    output = SerializeGetTypeExpression(tree);
                    break;

                case TreeType.CTypeExpression:
                case TreeType.DirectCastExpression:
                    output = SerializeCastTypeExpression(tree);
                    break;

                case TreeType.TypeOfExpression:
                    output = SerializeTypeOfExpression(tree);
                    break;

                case TreeType.IntrinsicCastExpression:
                    output = SerializeIntrinsicCastExpression(tree);
                    break;

                case TreeType.SimpleNameExpression:
                    output = SerializeSimpleNameExpression(tree);
                    break;

                case TreeType.QualifiedExpression:
                    output = SerializeQualifiedExpression(tree);
                    break;

                case TreeType.DictionaryLookupExpression:
                    output = SerializeDictionaryLookupExpression(tree);
                    break;

                case TreeType.InstanceExpression:
                    output = SerializeInstanceExpression(tree);
                    break;

                case TreeType.ParentheticalExpression:
                    output = SerializeParentheticalExpression(tree);
                    break;

                case TreeType.BinaryOperatorExpression:
                    output = SerializeBinaryOperatorExpression(tree);
                    break;

                case TreeType.UnaryOperatorExpression:
                    output = SerializeUnaryOperatorExpression(tree);
                    break;
                
                case TreeType.NewExpression:
                    output = SerializeNewExpression(tree);
                    break;

                case TreeType.NothingExpression:
                    output = SerializeNothingExpression(tree);
                    break;

                default:
                    throw new NotImplementedException("SerializeExpression not implemented for class " + tree.GetType().Name);
            }

            return output;
        }

        private CodeExpression SerializeCallOrIndexExpression(Tree tree)
        {
            if (tree == null)
                return null;

            CallOrIndexExpression input = (CallOrIndexExpression)tree;
            CodeExpression output = null;
            
            CodeExpression target = SerializeExpression(input.TargetExpression);
            CodeExpressionCollection arguments = SerializeArgumentCollection(input.Arguments);
            CodeExpression[] argumentArray = ConvertCodeExpressionCollection(arguments);

            if (input.IsIndex)
            {
                output = new CodeArrayIndexerExpression(target, argumentArray);
            }
            else
            {
                CodeMethodReferenceExpression method = null;
                if (input.TargetExpression is SimpleNameExpression)
                {
                    CodeExpression expression = SerializeSimpleNameExpression(input.TargetExpression);

                    if (expression is CodeMethodReferenceExpression)
                    {
                        method = (CodeMethodReferenceExpression)expression;
                    }
                    else if (target is CodeVariableReferenceExpression)
                    {
                        CodeVariableReferenceExpression variable = (CodeVariableReferenceExpression)expression;
                        method = new CodeMethodReferenceExpression(null, variable.VariableName);
                    }
                }
                else if (input.TargetExpression is QualifiedExpression)
                {
                    method = SerializeQualifiedExpression(input.TargetExpression);
                }
             
                CodeMethodReferenceExpression instrinsic = VisuallBasicUtil.MapIntrinsicName(method);
                if (instrinsic != null)
                    method = instrinsic;

                output = new CodeMethodInvokeExpression(method, argumentArray);
            }

            return output;            
        }

        private CodeExpression SerializeLiteralExpression(Tree tree)
        {
            if (tree == null)
                return null;

            Expression input = (Expression)tree;
            CodeExpression output = null;
            
            switch (input.Type)
            {
                case TreeType.BooleanLiteralExpression:
                    BooleanLiteralExpression booleanLiteral =(BooleanLiteralExpression)tree;
                    output = new CodePrimitiveExpression(booleanLiteral.Literal);
                    break;

                case TreeType.StringLiteralExpression:
                    StringLiteralExpression stringLiteral = (StringLiteralExpression)tree;
                    output = new CodePrimitiveExpression(stringLiteral.Literal);
                    break;

                case TreeType.CharacterLiteralExpression:
                    CharacterLiteralExpression characterLiteral = (CharacterLiteralExpression)tree;
                    output = new CodePrimitiveExpression(characterLiteral.Literal);
                    break;

                case TreeType.DateLiteralExpression:
                    DateLiteralExpression dateLiteral = (DateLiteralExpression)tree;
                    
                    DateTime date = dateLiteral.Literal;
                    List<CodeExpression> arguments = new List<CodeExpression>();
                    
                    arguments.Add(new CodePrimitiveExpression(date.Year));
                    arguments.Add(new CodePrimitiveExpression(date.Month));
                    arguments.Add(new CodePrimitiveExpression(date.Day));

                    if (date.Hour > 0 || date.Minute > 0 || date.Second > 0)
                    {
                        arguments.Add(new CodePrimitiveExpression(date.Hour));
                        arguments.Add(new CodePrimitiveExpression(date.Minute));
                        arguments.Add(new CodePrimitiveExpression(date.Second));
                    }

                    if (date.Millisecond > 0)
                        arguments.Add(new CodePrimitiveExpression(date.Millisecond));
                    
                    output = new CodeObjectCreateExpression(typeof(DateTime), arguments.ToArray());
                    break;

                case TreeType.IntegerLiteralExpression:
                    IntegerLiteralExpression integerLiteral = (IntegerLiteralExpression)tree;
                    output = new CodePrimitiveExpression(integerLiteral.Literal);
                    //integerLiteral.TypeCharacter
                    //integerLiteral.IntegerBase
                    break;

                case TreeType.FloatingPointLiteralExpression:
                    FloatingPointLiteralExpression floatLiteral = (FloatingPointLiteralExpression)tree;
                    output = new CodePrimitiveExpression(floatLiteral.Literal);
                    //floatLiteral.TypeCharacter
                    break;

                case TreeType.DecimalLiteralExpression:
                    DecimalLiteralExpression decimalLiteral = (DecimalLiteralExpression)tree;
                    output = new CodePrimitiveExpression(decimalLiteral.Literal);
                    //decimalLiteral.TypeCharacter
                    break;
            }

            return output;
        }

        private CodeExpression SerializeGetTypeExpression(Tree tree)
        {
            throw new NotImplementedException();
            //GetTypeExpression input = (GetTypeExpression)tree;
            //input.IsConstant
            //input.Target            
        }

        private CodeExpression SerializeCastTypeExpression(Tree tree)
        {
            throw new NotImplementedException();
            //CastTypeExpression input = (CastTypeExpression)tree;
            //input.Operand
            //input.Target
        }

        /// <summary>
        /// Represents the closest expression to IL's isinst opcode (C#'s is keyword).
        /// </summary>
        /// <example>Output example:<code>myVariable.GetType().IsInstanceOfType(typeof(System.IComparable))</code></example>
        private CodeMethodInvokeExpression SerializeTypeOfExpression(Tree tree)
        {
            if (tree == null)
                return null;

            TypeOfExpression input = (TypeOfExpression)tree;

            CodeExpression operand = SerializeExpression(input.Operand);
            CodeTypeReference target = SerializeTypeName(input.Target);
            
            CodeMethodInvokeExpression method = new CodeMethodInvokeExpression(operand, "GetType");
            CodeTypeOfExpression type = new CodeTypeOfExpression(target);
            CodeMethodInvokeExpression output = new CodeMethodInvokeExpression(method, "IsInstanceOfType", type);

            return output;            
        }

        private CodeExpression SerializeIntrinsicCastExpression(Tree tree)
        {
            if (tree == null)
                return null;

            IntrinsicCastExpression input = (IntrinsicCastExpression)tree;
            CodeExpression output = null;

            CodeExpression operand = SerializeExpression(input.Operand);
            Type targetType = TranslateIntrinsicType(input.IntrinsicType);
            CodeTypeReference target = new CodeTypeReference(targetType);

            switch (input.IntrinsicType)
            {
                case IntrinsicType.Boolean:
                case IntrinsicType.Byte:
                case IntrinsicType.Short:
                case IntrinsicType.Integer:
                case IntrinsicType.Long:
                case IntrinsicType.Decimal:
                case IntrinsicType.Single:
                case IntrinsicType.Double:
                case IntrinsicType.Date:
                case IntrinsicType.Currency:
                    CodeVariableReferenceExpression variable = new CodeVariableReferenceExpression(targetType.Name);
                    output = new CodeMethodInvokeExpression(variable, "Parse", operand);
                    break;
                case IntrinsicType.String:
                    output = new CodeMethodInvokeExpression(operand, "ToString");
                    break;
                case IntrinsicType.Object:
                case IntrinsicType.Variant:
                    output = new CodeCastExpression(target, operand);
                    break;
            }

            return output;
        }

        private CodeExpression SerializeSimpleNameExpression(Tree tree)
        {
            if (tree == null)
                return null;

            SimpleNameExpression input = (SimpleNameExpression)tree;
            CodeExpression output = null;

            string name = input.Name.Name;
            CodeVariableReferenceExpression variable = new CodeVariableReferenceExpression(name);
            output = variable;

            CodeMethodReferenceExpression instrinsic = VisuallBasicUtil.MapIntrinsicName(variable);
            if (instrinsic != null)
                output = instrinsic;
            
            return output;
        }

        private CodeMethodReferenceExpression SerializeQualifiedExpression(Tree tree)
        {
            if (tree == null)
                return null;

            QualifiedExpression input = (QualifiedExpression)tree;
            CodeExpression qualifier = SerializeExpression(input.Qualifier);
            string name = input.Name.Name;
            CodeMethodReferenceExpression output = new CodeMethodReferenceExpression(qualifier, name);

            CodeMethodReferenceExpression instrinsic = VisuallBasicUtil.MapIntrinsicName(output);
            if (instrinsic != null)
                output = instrinsic;

            return output;
        }

        private CodeIndexerExpression SerializeDictionaryLookupExpression(Tree tree)
        {
            if (tree == null)
                return null;

            DictionaryLookupExpression input = (DictionaryLookupExpression)tree;
            SimpleNameExpression nameExpression = new SimpleNameExpression(input.Name, input.Span);

            CodeExpression qualifier = SerializeExpression(input.Qualifier);
            CodeExpression name = null;

            if (LanguageInfo.ImplementsVB60(Version))
                name = new CodePrimitiveExpression(nameExpression.Name.Name);
            else
                name = SerializeExpression(nameExpression);

            CodeIndexerExpression output = new CodeIndexerExpression(qualifier, name);

            return output;
        }

        private CodeExpression SerializeInstanceExpression(Tree tree)
        {
            if (tree == null)
                return null;

            InstanceExpression input = (InstanceExpression)tree;
            CodeExpression output = null;

            switch (input.InstanceType)
            {
                case InstanceType.Me:
                    output = new CodeThisReferenceExpression();
                    break;
                case InstanceType.MyClass:
                    // TODO: Find out MyClass CodeDom equivalent (.NET)
                    output = new CodeThisReferenceExpression();
                    break;
                case InstanceType.MyBase:
                    output = new CodeBaseReferenceExpression();
                    break;
            }

            return output;
        }

        private CodeExpression SerializeParentheticalExpression(Tree tree)
        {
            if (tree == null)
                return null;

            ParentheticalExpression input = (ParentheticalExpression)tree;
            CodeExpression operand = SerializeExpression(input.Operand);

            CodeMethodInvokeExpression output = new CodeMethodInvokeExpression(null, operand);

            return output;
        }

        private CodeBinaryOperatorExpression SerializeUnaryOperatorExpression(Tree tree)
        {
            if (tree == null)
                return null;

            UnaryOperatorExpression input = (UnaryOperatorExpression)tree;
            CodeBinaryOperatorExpression output = null;

            CodeExpression operand = SerializeExpression(input.Operand);
            CodeExpression auxiliar = null;

            switch (input.Operator)
            {
                case UnaryOperatorType.UnaryPlus:
                    auxiliar = new CodePrimitiveExpression(0);
                    output = new CodeBinaryOperatorExpression(auxiliar, CodeBinaryOperatorType.Add, operand);
                    break;
                case UnaryOperatorType.Negate:
                    auxiliar = new CodePrimitiveExpression(-1);
                    output = new CodeBinaryOperatorExpression(auxiliar, CodeBinaryOperatorType.Multiply, operand);
                    break;
                case UnaryOperatorType.Not:
                    auxiliar = new CodePrimitiveExpression(false);
                    output = new CodeBinaryOperatorExpression(operand, CodeBinaryOperatorType.ValueEquality, auxiliar);
                    break;
            }

            return output;
        }

        private CodeExpression SerializeBinaryOperatorExpression(Tree tree)
        {
            if (tree == null)
                return null;

            BinaryOperatorExpression input = (BinaryOperatorExpression)tree;
            CodeExpression output = null;

            CodeExpression arraySize = GetExplicitArraySize(input);
            if (arraySize != null)
                return arraySize;

            CodeExpression left = SerializeExpression(input.LeftOperand);
            CodeExpression right = SerializeExpression(input.RightOperand);

            switch (input.Operator)
            {
                case BinaryOperatorType.Plus:
                case BinaryOperatorType.Concatenate:
                    output = new CodeBinaryOperatorExpression(left, CodeBinaryOperatorType.Add, right);
                    break;
                case BinaryOperatorType.Minus:
                    output = new CodeBinaryOperatorExpression(left, CodeBinaryOperatorType.Subtract, right);
                    break;
                case BinaryOperatorType.Multiply:
                    output = new CodeBinaryOperatorExpression(left, CodeBinaryOperatorType.Multiply, right);
                    break;
                case BinaryOperatorType.Divide:
                    output = new CodeBinaryOperatorExpression(left, CodeBinaryOperatorType.Divide, right);
                    break;
                case BinaryOperatorType.IntegralDivide:
                    CodeBinaryOperatorExpression divide = new CodeBinaryOperatorExpression(left, CodeBinaryOperatorType.Divide, right);
                    break;
                case BinaryOperatorType.Power:
                    CodeMethodReferenceExpression method = new CodeMethodReferenceExpression(new CodeTypeReferenceExpression("Math"), "Pow");
                    output = new CodeMethodInvokeExpression(method, left, right);
                    break;
                case BinaryOperatorType.Modulus:
                    output = new CodeBinaryOperatorExpression(left, CodeBinaryOperatorType.Modulus, right);
                    break;
                case BinaryOperatorType.Or:
                    output = new CodeBinaryOperatorExpression(left, CodeBinaryOperatorType.BooleanOr, right);
                    break;
                case BinaryOperatorType.And:
                    output = new CodeBinaryOperatorExpression(left, CodeBinaryOperatorType.BooleanAnd, right);
                    break;
                case BinaryOperatorType.Is:
                    output = new CodeBinaryOperatorExpression(left, CodeBinaryOperatorType.IdentityEquality, right);
                    break;
                case BinaryOperatorType.IsNot:
                    output = new CodeBinaryOperatorExpression(left, CodeBinaryOperatorType.IdentityInequality, right);
                    break;
                case BinaryOperatorType.Equals:
                    output = new CodeBinaryOperatorExpression(left, CodeBinaryOperatorType.ValueEquality, right);
                    break;
                case BinaryOperatorType.NotEquals:
                    CodeBinaryOperatorExpression equals = new CodeBinaryOperatorExpression(left, CodeBinaryOperatorType.ValueEquality, right);
                    output = new CodeBinaryOperatorExpression(equals, CodeBinaryOperatorType.ValueEquality, new CodePrimitiveExpression(false));
                    break;
                case BinaryOperatorType.LessThan:
                    output = new CodeBinaryOperatorExpression(left, CodeBinaryOperatorType.LessThan, right);
                    break;
                case BinaryOperatorType.LessThanEquals:
                    output = new CodeBinaryOperatorExpression(left, CodeBinaryOperatorType.LessThanOrEqual, right);
                    break;
                case BinaryOperatorType.GreaterThan:
                    output = new CodeBinaryOperatorExpression(left, CodeBinaryOperatorType.GreaterThan, right);
                    break;
                case BinaryOperatorType.GreaterThanEquals:
                    output = new CodeBinaryOperatorExpression(left, CodeBinaryOperatorType.GreaterThanOrEqual, right);
                    break;
                case BinaryOperatorType.To:
                    CodeVariableReferenceExpression temp = new CodeVariableReferenceExpression("_to_operator");
                    CodeBinaryOperatorExpression minor = new CodeBinaryOperatorExpression(temp, CodeBinaryOperatorType.GreaterThanOrEqual, left);
                    CodeBinaryOperatorExpression major = new CodeBinaryOperatorExpression(temp, CodeBinaryOperatorType.LessThanOrEqual, right);
                    output = new CodeBinaryOperatorExpression(minor, CodeBinaryOperatorType.BooleanAnd, major);
                    break;
                case BinaryOperatorType.Like:
                    bool compareText = HasOptionType(OptionType.CompareText);
                    string compareMethod = compareText ? "CompareMethod.Text" : "CompareMethod.Binary";
                    CodeMethodReferenceExpression compareExpression = new CodeMethodReferenceExpression(new CodeTypeReferenceExpression("LikeOperator"), "LikeString");
                    CodeTypeReferenceExpression compareType = new CodeTypeReferenceExpression(new CodeTypeReference(compareMethod));
                    output = new CodeMethodInvokeExpression(compareExpression, left, right, compareType);
                    break;
                case BinaryOperatorType.ShiftLeft:
                    throw new NotImplementedException("'ShiftLeft' operator not implemented");
                case BinaryOperatorType.ShiftRight:
                    throw new NotImplementedException("'ShiftRight' operator not implemented");
                case BinaryOperatorType.OrElse:
                    throw new NotImplementedException("'OrElse' operator not implemented");
                case BinaryOperatorType.AndAlso:
                    throw new NotImplementedException("'AndAlso' operator not implemented");
                case BinaryOperatorType.Xor:
                    throw new NotImplementedException("'Xor' operator not implemented");
                default:
                    throw new NotImplementedException("SerializeBinaryOperator not implemented for class " + input.GetType().Name);
            }

            return output;
        }

        private CodeExpression SerializeNewExpression(Tree tree)
        {
            if (tree == null)
                return null;

            NewExpression input = (NewExpression)tree;

            //input.Arguments
            CodeTypeReference type = SerializeTypeName(input.Target);
            CodeObjectCreateExpression output = new CodeObjectCreateExpression(type);

            return output;
        }

        private CodePrimitiveExpression SerializeNothingExpression(Tree tree)
        {
            if (tree == null)
                return null;

            NothingExpression input = (NothingExpression)tree;
            CodePrimitiveExpression output = new CodePrimitiveExpression(null);
            return output;
        }

        #endregion

        #region Statements

        private CodeStatementCollection SerializeStatement(Tree tree)
        {
            if (tree == null)
                return null;

            CodeStatementCollection output = null;

            switch (tree.Type)
            {
                case TreeType.GotoStatement:
                    output = SerializeGotoStatement(tree);
                    break;
                
                case TreeType.GoSubStatement:
                    output = SerializeGoSubStatement(tree);
                    break;

                case TreeType.LabelStatement:
                    output = SerializeLabelStatement(tree);
                    break;

                case TreeType.ContinueStatement:
                    output = SerializeContinueStatement(tree);
                    break;

                case TreeType.ExitStatement:
                    output = SerializeExitStatement(tree);
                    break;

                case TreeType.ReturnStatement:
                    output = SerializeReturnStatement(tree);
                    break;

                case TreeType.ErrorStatement:
                    output = SerializeErrorStatement(tree);
                    break;

                case TreeType.ThrowStatement:
                    output = SerializeThrowStatement(tree);
                    break;

                case TreeType.RaiseEventStatement:
                    output = SerializeRaiseEventStatement(tree);
                    break;

                case TreeType.AddHandlerStatement:
                case TreeType.RemoveHandlerStatement:
                    output = SerializeHandlerStatement(tree);
                    break;

                case TreeType.OnErrorStatement:
                    output = SerializeOnErrorStatement(tree);
                    break;

                case TreeType.ResumeStatement:
                    output = SerializeResumeStatement(tree);
                    break;

                case TreeType.ReDimStatement:
                    output = SerializeReDimStatement(tree);
                    break;

                case TreeType.EraseStatement:
                    output = SerializeEraseStatement(tree);
                    break;

                case TreeType.CallStatement:
                    output = SerializeCallStatement(tree);
                    break;

                case TreeType.AssignmentStatement:
                    output = SerializeAssignmentStatement(tree);
                    break;

                case TreeType.MidAssignmentStatement:
                    output = SerializeMidAssignmentStatement(tree);
                    break;

                case TreeType.CompoundAssignmentStatement:
                    output = SerializeCompoundAssignmentStatement(tree);
                    break;

                case TreeType.LocalDeclarationStatement:
                    output = SerializeLocalDeclarationStatement(tree);
                    break;

                case TreeType.EndBlockStatement:
                    output = SerializeEndBlockStatement(tree);
                    break;

                case TreeType.WhileBlockStatement:
                    output = SerializeWhileBlockStatement(tree);
                    break;

                case TreeType.WithBlockStatement:
                    output = SerializeWithBlockStatement(tree);
                    break;

                case TreeType.SyncLockBlockStatement:
                    output = SerializeSyncLockBlockStatement(tree);
                    break;

                case TreeType.UsingBlockStatement:
                    output = SerializeUsingBlockStatement(tree);
                    break;

                case TreeType.DoBlockStatement:
                    output = SerializeDoBlockStatement(tree);
                    break;

                case TreeType.LoopStatement:
                    output = SerializeLoopStatement(tree);
                    break;

                case TreeType.NextStatement:
                    output = SerializeNextStatement(tree);
                    break;

                case TreeType.ForBlockStatement:
                    output = SerializeForBlockStatement(tree);
                    break;

                case TreeType.ForEachBlockStatement:
                    output = SerializeForEachBlockStatement(tree);
                    break;

                case TreeType.CatchStatement:
                    output = SerializeCatchStatement(tree);
                    break;
                /*
                case TreeType.CaseStatement:
                    output = SerializeCaseStatement(tree);
                    break;
                */
                case TreeType.CaseElseStatement:
                    output = SerializeCaseElseStatement(tree);
                    break;

                case TreeType.CaseBlockStatement:
                    output = SerializeCaseBlockStatement(tree);
                    break;

                case TreeType.CaseElseBlockStatement:
                    output = SerializeCaseElseBlockStatement(tree);
                    break; 

                case TreeType.SelectBlockStatement:
                    output = SerializeSelectBlockStatement(tree);
                    break;

                case TreeType.IfBlockStatement:
                    output = SerializeIfBlockStatement(tree);
                    break;

                case TreeType.ElseBlockStatement:
                    output = SerializeElseBlockStatement(tree);
                    break;

                case TreeType.ElseIfStatement:
                    output = SerializeElseIfStatement(tree);
                    break;

                case TreeType.ElseIfBlockStatement:
                    output = SerializeElseIfBlockStatement(tree);
                    break;

                case TreeType.LineIfBlockStatement:
                    output = SerializeLineIfBlockStatement(tree);
                    break;

                case TreeType.EmptyStatement:
                    output = SerializeEmptyStatement(tree);
                    break;

                case TreeType.EndStatement:
                    output = SerializeEndStatement(tree);
                    break;

                default:
                    throw new NotImplementedException("SerializeStatement not implemented for class " + tree.GetType().Name);
            }

            return output;
        }

        private CodeStatementCollection SerializeGotoStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            GotoStatement input = (GotoStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);

            CodeReferenceName reference = SerializeName(input.Name);
            CodeGotoStatement output = new CodeGotoStatement(reference.Name);
            
            list.Add(output);
            return list;
        }

        private CodeStatementCollection SerializeGoSubStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            GoSubStatement input = (GoSubStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);
            
            CodeReferenceName reference = SerializeName(input.Name);
            CodeGotoStatement output = new CodeGotoStatement(reference.Name);
            list.Add(output);

            CodeLabeledStatement returnLabel = new CodeLabeledStatement(input.ReturnLabel);
            list.Add(returnLabel);

            return list;
        }

        private CodeStatementCollection SerializeLabelStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            LabelStatement input = (LabelStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);

            CodeReferenceName reference = SerializeName(input.Name);
            CodeLabeledStatement output = new CodeLabeledStatement(reference.Name);

            list.Add(output);
            return list;
        }

        private CodeStatementCollection SerializeContinueStatement(Tree tree)
        {
            throw new NotImplementedException();
            //ContinueStatement input = (ContinueStatement)tree;
            //input.Comments
            //input.ContinueType
        }

        private CodeStatementCollection SerializeExitStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            ExitStatement input = (ExitStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);
            
            CodeStatement output = null;
            
            switch (input.ExitType)
            {
                case BlockType.Do:
                    output = new CodeBreakStatement();
                    break;
                
                case BlockType.For:
                    output = new CodeBreakStatement();
                    break;
                
                case BlockType.While:
                    output = new CodeBreakStatement();
                    break;
                
                case BlockType.Sub:
                    output = new CodeMethodReturnStatement();
                    break;
                
                case BlockType.Function:
                    CodeTypeReference functionType = null;
                    if (input.RelatedTree is FunctionDeclaration)
                    {
                        FunctionDeclaration function = (FunctionDeclaration)input.RelatedTree;
                        functionType = SerializeTypeName(function.ResultType);
                    }

                    CodeDefaultValueExpression functionValue = new CodeDefaultValueExpression(functionType);
                    output = new CodeMethodReturnStatement(functionValue);
                    break;

                case BlockType.Property:
                    CodeTypeReference propertyType = null;
                    CodeDefaultValueExpression propertyValue = null;
                    
                    if (input.RelatedTree is GetAccessorDeclaration)
                    {
                        GetAccessorDeclaration accessor = (GetAccessorDeclaration)input.RelatedTree;
                        PropertyDeclaration property = null;
                        if (accessor.Parent is PropertyDeclaration)
                            property = (PropertyDeclaration)accessor.Parent;

                        if (property != null)
                        {
                            propertyType = SerializeTypeName(property.ResultType);
                            propertyValue = new CodeDefaultValueExpression(propertyType);
                        }
                    }
                    
                    output = new CodeMethodReturnStatement(propertyValue);
                    break;

                default:
                    output = new CodeCommentStatement(string.Format("TODO: ExitStatement Not Supported ({0})", input.ExitType));
                    break;
            }

            list.Add(output);
            return list;
        }

        private CodeStatementCollection SerializeReturnStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            ReturnStatement input = (ReturnStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);

            CodeStatement output = null;
            
            CodeExpression expression = SerializeExpression(input.Expression);
            string labelName = string.Empty;

            if (input.IsLabelReference && input.GoSubReference != null)
                labelName = input.GoSubReference.ReturnLabel;

            if (LanguageInfo.ImplementsVB60(Version))
            {
                if (!string.IsNullOrEmpty(labelName))
                    output = new CodeGotoStatement(labelName);
                else
                    output = new CodeCommentStatement("There is no GoSub Statement for this Return Statement.");
            }
            else
            {
                output = new CodeMethodReturnStatement(expression);
            }

            list.Add(output);
            list.Add(new CodeConditionStatement(new CodePrimitiveExpression(true)));
            return list;
        }

        private CodeStatementCollection SerializeErrorStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            ErrorStatement input = (ErrorStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);

            CodeObjectCreateExpression create = new CodeObjectCreateExpression("ErrObject");
            CodeVariableDeclarationStatement err = new CodeVariableDeclarationStatement("ErrObject", "Err", create);
            list.Add(err);

            CodeVariableReferenceExpression variable = new CodeVariableReferenceExpression("Err");
            CodeMethodReferenceExpression method = new CodeMethodReferenceExpression(variable, "Raise");

            CodeExpression expression = SerializeExpression(input.Expression);
            CodeMethodInvokeExpression output = new CodeMethodInvokeExpression(method, expression);
            
            list.Add(output);
            return list;
        }

        private CodeStatementCollection SerializeThrowStatement(Tree tree)
        {
            throw new NotImplementedException();
            //ThrowStatement input = (ThrowStatement)tree;
            //input.Comments
            //input.Expression
        }

        private CodeStatementCollection SerializeRaiseEventStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            RaiseEventStatement input = (RaiseEventStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);

            CodeMethodInvokeExpression output = new CodeMethodInvokeExpression();
            
            CodeReferenceName reference = SerializeName(input.Name);
            CodeExpressionCollection arguments = SerializeArgumentCollection(input.Arguments);
            
            CodeMethodReferenceExpression method = new CodeMethodReferenceExpression(null, reference.Name);
            
            output.Method = method;
            output.Parameters.AddRange(arguments);

            list.Add(output);
            return list;
        }

        private CodeStatementCollection SerializeHandlerStatement(Tree tree)
        {
            throw new NotImplementedException();
            //HandlerStatement input = (HandlerStatement)tree;
            //input.Comments
            //input.Name
            //input.DelegateExpression
        }

        private CodeStatementCollection SerializeOnErrorStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            OnErrorStatement input = (OnErrorStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);

            CodeStatement output = null;
            
            switch (input.OnErrorType)
            {
                case OnErrorType.Zero:
                    output = new CodeCommentStatement("TODO: 'On Error GoTo 0' Statement not supported");
                    break;

                case OnErrorType.MinusOne:
                    output = new CodeCommentStatement("TODO: 'On Error GoTo -1' Statement not supported");
                    break;

                case OnErrorType.Label:
                    output = new CodeCommentStatement("TODO: 'On Error GoTo [Label]' Statement not supported");
                    break;

                case OnErrorType.Next:
                    output = new CodeCommentStatement("TODO: 'On Error Resume' Next Statement not supported");
                    break;

                default:
                    throw new NotImplementedException("SerializeOnErrorStatement not implemented for the error type " + input.OnErrorType.ToString());
            }

            list.Add(output);
            return list;
        }

        private CodeStatementCollection SerializeResumeStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            ResumeStatement input = (ResumeStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);

            CodeStatement output = null;

            switch (input.ResumeType)
            {
                case ResumeType.Next:
                    output = new CodeCommentStatement("TODO: 'Resume Next' Statement not supported");
                    break;

                case ResumeType.Label:
                    CodeReferenceName label = SerializeName(input.Name);
                    output = new CodeGotoStatement(label.Name);
                    break;

                case ResumeType.None:
                    output = new CodeCommentStatement("TODO: 'Resume' Statement not supported");
                    break;

                default:
                    throw new NotImplementedException("SerializeResumeStatement not implemented for the error type " + input.ResumeType.ToString());
            }

            list.Add(output);
            return list;
        }

        private CodeStatementCollection SerializeReDimStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            ReDimStatement input = (ReDimStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection comentStatements = ConvertComments(comments);
            list.AddRange(comentStatements);

            foreach (Expression variable in input.Variables)
            {
                if (variable is CallOrIndexExpression)
                {
                    CallOrIndexExpression index = (CallOrIndexExpression)variable;
                    string tempName = "array";

                    if (index.TargetExpression is SimpleNameExpression)
                        tempName = ((SimpleNameExpression)index.TargetExpression).Name.Name;
                    else if (index.TargetExpression is QualifiedExpression)
                        tempName = ((QualifiedExpression)index.TargetExpression).Name.Name;

                    tempName = "_temp_" + tempName;

                    CodeExpression target = SerializeExpression(index.TargetExpression);

                    CodeTypeReference type = new CodeTypeReference(typeof(object));

                    if (index.Declarator != null)
                        type = SerializeTypeName(index.Declarator.VariableType);

                    CodeExpressionCollection arguments = SerializeArgumentCollection(index.Arguments);
                    CodeExpression[] argumentArray = ConvertCodeExpressionCollection(arguments);
                    int rank = arguments == null ? 0 : arguments.Count;
                    type.ArrayRank = rank;

                    CodeExpression init = CreateArrayInitializer(type, arguments);
                    CodeExpression value = null;
                    
                    if (input.IsPreserve)
                    {
                        // Declare the temp array
                        CodeVariableDeclarationStatement declare = new CodeVariableDeclarationStatement(type, tempName, init);

                        CodeVariableReferenceExpression array = new CodeVariableReferenceExpression("Array");
                        CodeVariableReferenceExpression temp = new CodeVariableReferenceExpression(tempName);
                        CodePropertyReferenceExpression length = new CodePropertyReferenceExpression(target, "Length");
                        
                        CodeMethodInvokeExpression copy = new CodeMethodInvokeExpression(array, "Copy", target, temp, length);
                        
                        list.Add(declare);
                        list.Add(copy);

                        value = temp;
                    }
                    else
                    {
                        value = init;
                    }
                    
                    CodeAssignStatement assign = new CodeAssignStatement(target, value);
                    list.Add(assign);
                }
                else
                {
                    throw new InvalidOperationException("Invalid expression in ReDimStatement: " + variable.GetType().Name);
                }
            }

            return list;
        }

        private CodeStatementCollection SerializeEraseStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            EraseStatement input = (EraseStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);

            foreach (Expression item in input.Variables)
            {
                CodeExpression expression = SerializeExpression(item);
                CodePrimitiveExpression zero = new CodePrimitiveExpression(0);
                CodePropertyReferenceExpression length = new CodePropertyReferenceExpression(expression, "Length");
                CodeMethodInvokeExpression method = new CodeMethodInvokeExpression(expression, "Clear", zero, length);

                list.Add(method);
            }
            
            return list;
        }

        private CodeStatementCollection SerializeCallStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            CallStatement input = (CallStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);

            CodeExpressionCollection arguments = SerializeArgumentCollection(input.Arguments);
            CodeExpression[] argumentList = ConvertCodeExpressionCollection(arguments);
            
            CodeExpression target = null;
            string name = string.Empty;

            if (input.TargetExpression is SimpleNameExpression)
            {
                SimpleNameExpression simple = (SimpleNameExpression)input.TargetExpression;
                name = simple.Name.Name;
            }
            else if (input.TargetExpression is QualifiedExpression)
            {
                QualifiedExpression qualified = (QualifiedExpression)input.TargetExpression;
                target = SerializeExpression(qualified.Qualifier);
                name = qualified.Name.Name;
            }

            CodeMethodReferenceExpression method = new CodeMethodReferenceExpression(target, name);
            CodeMethodReferenceExpression instrinsic = VisuallBasicUtil.MapIntrinsicName(method);
            if (instrinsic != null)
                method = instrinsic;

            CodeMethodInvokeExpression output = new CodeMethodInvokeExpression(method, argumentList);
            list.Add(output);
            
            return list;
        }

        private CodeStatementCollection SerializeAssignmentStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            AssignmentStatement input = (AssignmentStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);

            CodeExpression target = SerializeExpression(input.TargetExpression);
            CodeExpression source = SerializeExpression(input.SourceExpression);
            CodeAssignStatement output = new CodeAssignStatement(target, source);

            list.Add(output);
            return list;
        }

        private CodeStatementCollection SerializeMidAssignmentStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            //Instructions: Replaces a specified number of characters in a String variable with characters from another string.
            MidAssignmentStatement input = (MidAssignmentStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);

            //input.HasTypeCharacter
            
            CodeExpression source = SerializeExpression(input.SourceExpression);
            CodeExpression target = SerializeExpression(input.TargetExpression);
            
            CodeExpression start = SerializeExpression(input.StartExpression);
            CodeExpression length = SerializeExpression(input.LengthExpression);

            CodeMethodReferenceExpression method = new CodeMethodReferenceExpression(new CodeVariableReferenceExpression("Strings"), "Mid");
            CodeMethodInvokeExpression invoke = new CodeMethodInvokeExpression(method, source, start, length);
            CodeAssignStatement output = new CodeAssignStatement(invoke, method);

            list.Add(output);
            return list;
        }

        private CodeStatementCollection SerializeCompoundAssignmentStatement(Tree tree)
        {
            throw new NotImplementedException();
            //CompoundAssignmentStatement input = (CompoundAssignmentStatement)tree;
            //input.Comments
            //input.TargetExpression
            //input.CompoundOperator
            //input.SourceExpression
        }

        private CodeStatementCollection SerializeLocalDeclarationStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            LocalDeclarationStatement input = (LocalDeclarationStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);
            
            // Local variables doesn't have modifiers
            MemberAttributes modifiers = SerializeModifierCollection(input.Modifiers);
            CodeVariableDeclaratorCollection declarations = SerializeVariableDeclaratorCollection(input.VariableDeclarators);

            if (declarations.CollectionType == CodeCollectionType.Statement)
                list.AddRange(declarations.Statements);
            else
                throw new Exception("Invalid collection in SerializeLocalDeclarationStatement: " + declarations.GetType().Name);
            
            return list;
        }

        private CodeStatementCollection SerializeEndBlockStatement(Tree tree)
        {
            throw new NotImplementedException();
            //EndBlockStatement input = (EndBlockStatement)tree;
            //input.Comments
            //input.EndType
        }

        private CodeStatementCollection SerializeWhileBlockStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            WhileBlockStatement input = (WhileBlockStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);

            CodeExpression condition = SerializeExpression(input.Expression);
            CodeStatementCollection statements = SerializeStatementCollection(input.Statements);
            
            CodeIterationStatement output = new CodeIterationStatement();
            output.InitStatement = new CodeSnippetStatement("");
            output.IncrementStatement = new CodeSnippetStatement("");
            output.TestExpression = condition;
            output.Statements.AddRange(statements);

            list.Add(output);
            return list;
        }

        private CodeStatementCollection SerializeWithBlockStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            WithBlockStatement input = (WithBlockStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);

            CodeStatementCollection statements = SerializeStatementCollection(input.QualifiedStatements);

            list.AddRange(statements);
            return list;
        }

        private CodeStatementCollection SerializeSyncLockBlockStatement(Tree tree)
        {
            throw new NotImplementedException();
            //SyncLockBlockStatement input = (SyncLockBlockStatement)tree;
            //input.Comments
            //input.Expression
            //input.Statements
        }

        private CodeStatementCollection SerializeUsingBlockStatement(Tree tree)
        {
            throw new NotImplementedException();
            //UsingBlockStatement input = (UsingBlockStatement)tree;
            //input.Comments
            //if (input.Expression != null)
            //    input.Expression
            //else
            //    input.VariableDeclarators
            //input.Statements
            //input.EndStatement
        }

        private CodeStatementCollection SerializeDoBlockStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            DoBlockStatement input = (DoBlockStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);

            bool isWhile = true;
            CodeExpression condition = new CodePrimitiveExpression(true);
            CodeStatement init = new CodeSnippetStatement("");
            CodeStatement increment = new CodeSnippetStatement("");
            
            CodeStatementCollection statements = SerializeStatementCollection(input.Statements);

            if (input.Expression != null)
            {
                isWhile = input.IsWhile;
                
                condition = SerializeExpression(input.Expression);
                init = new CodeSnippetStatement("");
                increment = new CodeSnippetStatement("");
            }
            else if (input.EndStatement != null && input.EndStatement.Expression != null)
            {
                CodeExpression expression = SerializeExpression(input.EndStatement.Expression);
                CodeVariableReferenceExpression variable = new CodeVariableReferenceExpression("_iterator");
                isWhile = input.EndStatement.IsWhile;

                condition = variable;
                init = new CodeVariableDeclarationStatement(typeof(bool), "_iterator", new CodePrimitiveExpression(true));
                increment = new CodeAssignStatement(variable, expression);
            }
            
            if (!isWhile)
                condition = new CodeBinaryOperatorExpression(condition, CodeBinaryOperatorType.ValueEquality, new CodePrimitiveExpression(false));
            
            CodeIterationStatement output = new CodeIterationStatement();
            output.InitStatement = init;
            output.TestExpression = condition;
            output.IncrementStatement = increment;
            output.Statements.AddRange(statements);

            list.Add(output);
            return list;
        }

        private CodeStatementCollection SerializeLoopStatement(Tree tree)
        {
            throw new NotImplementedException();
            //LoopStatement input = (LoopStatement)tree;
            //input.Comments
            //input.IsWhile
            //input.Expression
        }

        private CodeStatementCollection SerializeNextStatement(Tree tree)
        {
            throw new NotImplementedException();
            //NextStatement input = (NextStatement)tree;
            //input.Comments
            //input.Variables
        }

        private CodeStatementCollection SerializeForBlockStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            ForBlockStatement input = (ForBlockStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);

            //input.ControlVariableDeclarator

            CodeExpression expression = SerializeExpression(input.ControlExpression);
            CodeExpression lower = SerializeExpression(input.LowerBoundExpression);
            CodeExpression upper = SerializeExpression(input.UpperBoundExpression);
            CodeExpression step = SerializeExpression(input.StepExpression);

            if (input.StepExpression == null || step == null)
                step = new CodePrimitiveExpression(1);

            CodeAssignStatement init = new CodeAssignStatement(expression, lower);
            CodeBinaryOperatorExpression condition = new CodeBinaryOperatorExpression(expression, CodeBinaryOperatorType.LessThanOrEqual, upper);

            CodeBinaryOperatorExpression add = new CodeBinaryOperatorExpression(expression, CodeBinaryOperatorType.Add, step);
            CodeAssignStatement increment = new CodeAssignStatement(expression, add);

            CodeStatementCollection statements = SerializeStatementCollection(input.Statements);

            CodeIterationStatement output = new CodeIterationStatement();
            output.TestExpression = condition;
            output.InitStatement = init;
            output.IncrementStatement = increment;
            output.Statements.AddRange(statements);

            list.Add(output);
            return list;
        }

        private CodeStatementCollection SerializeForEachBlockStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            ForEachBlockStatement input = (ForEachBlockStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);

            CodeExpression collection = SerializeExpression(input.CollectionExpression);
            CodeMethodInvokeExpression invoke = new CodeMethodInvokeExpression(collection, "GetEnumerator");
            CodeVariableReferenceExpression enumerator = new CodeVariableReferenceExpression("_enumerator");

            CodeMethodInvokeExpression condition = new CodeMethodInvokeExpression(enumerator, "MoveNext");
            CodeVariableDeclarationStatement init = new CodeVariableDeclarationStatement("IEnumerator", "_enumerator", invoke);
            CodeSnippetStatement increment = new CodeSnippetStatement("");
            CodeStatementCollection statements = SerializeStatementCollection(input.Statements);

            CodeVariableDeclaratorCollection declarators = SerializeVariableDeclarator(input.ControlVariableDeclarator);
            CodeExpression control = SerializeExpression(input.ControlExpression);

            if (input.ControlExpression == null || !(input.ControlExpression is SimpleNameExpression))
                throw new InvalidOperationException("The control expression is not a SimpleNameExpression");
            
            if (!(control is CodeVariableReferenceExpression))
	            throw new InvalidOperationException("The converted control expression is not a CodeVariableReferenceExpression");

            CodeVariableReferenceExpression reference = (CodeVariableReferenceExpression)control;
            CodeTypeReference type = new CodeTypeReference(typeof(object));

            switch (declarators.CollectionType)
            {
                case CodeCollectionType.TypeMember:
                    //CodeMemberField
                    foreach (CodeTypeMember item in declarators.Members)
                    {
                        if (item is CodeMemberField)
                        {
                            CodeMemberField member = (CodeMemberField)item;
                            if (member.Name == reference.VariableName)
                            {
                                type = member.Type;
                                break;
                            }
                        }
                    }
                    break;
                case CodeCollectionType.Statement:
                    //CodeVariableDeclarationStatement
                    foreach (CodeStatement item in declarators.Statements)
                    {
                        if (item is CodeVariableDeclarationStatement)
                        {
                            CodeVariableDeclarationStatement declaration = (CodeVariableDeclarationStatement)item;
                            if (declaration.Name == reference.VariableName)
                            {
                                type = declaration.Type;
                                break;
                            }
                        }
                    }
                    break;
            }
            
            //int item = ((int)(_enumerator.Current));
            CodePropertyReferenceExpression property = new CodePropertyReferenceExpression(enumerator, "Current");
            CodeCastExpression cast = new CodeCastExpression(type, property);
            CodeAssignStatement controlItem = new CodeAssignStatement(control, cast);

            CodeIterationStatement output = new CodeIterationStatement();
            output.TestExpression = condition;
            output.InitStatement = init;
            output.IncrementStatement = increment;

            output.Statements.Add(controlItem);
            output.Statements.AddRange(statements);

            list.Add(output);
            return list;
        }

        private CodeStatementCollection SerializeCatchStatement(Tree tree)
        {
            throw new NotImplementedException();
            //CatchStatement input = (CatchStatement)tree;
            //input.Comments
            //input.Name
            //input.AsLocation
            //input.ExceptionType
            //input.WhenLocation
            //input.FilterExpression
        }

        private CodeExpressionCollection SerializeCaseStatement(Tree tree)
        {
            if (tree == null)
                return new CodeExpressionCollection();

            CaseStatement input = (CaseStatement)tree;

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeExpressionCollection output = SerializeCaseClauseCollection(input.CaseClauses);
            
            return output;
        }

        private CodeStatementCollection SerializeCaseElseStatement(Tree tree)
        {
            throw new NotImplementedException();
            //CaseElseStatement input = (CaseElseStatement)tree;
            //input.Comments
            //input.ElseLocation
        }


        private CodeStatementCollection SerializeCaseBlockStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            CaseBlockStatement input = (CaseBlockStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);
            
            CodeExpressionCollection conditions = SerializeCaseStatement(input.CaseStatement);
            CodeStatementCollection statements = SerializeStatementCollection(input.Statements);

            CodeExpression condition = null;

            foreach (CodeExpression item in conditions)
            {
                if (condition == null)
                    condition = item;
                else
                    condition = new CodeBinaryOperatorExpression(condition, CodeBinaryOperatorType.BooleanAnd, item);
            }
            
            CodeConditionStatement output = new CodeConditionStatement();
            output.Condition = condition;
            output.TrueStatements.AddRange(statements);

            list.Add(output);
            return list;
        }

        private CodeStatementCollection SerializeCaseElseBlockStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            CaseElseBlockStatement input = (CaseElseBlockStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);

            //input.CaseElseStatement
            CodeStatementCollection statements = SerializeStatementCollection(input.Statements);

            list.AddRange(statements);
            return list;
        }

        private CodeStatementCollection SerializeSelectBlockStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            SelectBlockStatement input = (SelectBlockStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);

            //input.Expression

            CodeStatementCollection caseStatements = SerializeStatementCollection(input.CaseBlockStatements);
            CodeStatementCollection elseStatements = SerializeCaseElseBlockStatement(input.CaseElseBlockStatement);

            CodeConditionStatement output = null;
            CodeConditionStatement current = null;

            foreach (CodeStatement item in caseStatements)
            {
                if (!(item is CodeConditionStatement))
                    throw new InvalidOperationException("Invalid CodeStatement: " + item.GetType().Name);

                if (current == null)
                    output = (CodeConditionStatement)item;
                else
                    current.FalseStatements.Add(item);
                
                current = (CodeConditionStatement)item;
            }

            if (elseStatements != null && current != null)
                current.FalseStatements.AddRange(elseStatements);
            
            list.Add(output);
            return list;
        }

        private CodeStatementCollection SerializeIfBlockStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection(); 
            
            IfBlockStatement input = (IfBlockStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);

            CodeExpression condition = SerializeExpression(input.Expression);
            CodeStatementCollection trueStatements = SerializeStatementCollection(input.Statements);
            CodeStatementCollection elseIfStatements = SerializeStatementCollection(input.ElseIfBlockStatements);
            CodeStatementCollection elseStatements = SerializeElseBlockStatement(input.ElseBlockStatement);

            CodeConditionStatement output = new CodeConditionStatement();

            output.Condition = condition;
            output.TrueStatements.AddRange(trueStatements);

            CodeConditionStatement current = output;

            foreach (CodeStatement item in elseIfStatements)
            {
                if (!(item is CodeConditionStatement))
                    throw new InvalidOperationException("Invalid CodeStatement: " + item.GetType().Name);

                if (current == null)
                    output = (CodeConditionStatement)item;
                else
                    current.FalseStatements.Add(item);

                current = (CodeConditionStatement)item;
            }

            current.FalseStatements.AddRange(elseStatements);

            list.Add(output);
            return list;
        }

        private CodeStatementCollection SerializeElseBlockStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            ElseBlockStatement input = (ElseBlockStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);

            CodeStatementCollection statements = SerializeStatementCollection(input.Statements);

            list.AddRange(statements);
            return list;
        }

        private CodeStatementCollection SerializeElseIfStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            ElseIfStatement input = (ElseIfStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);
            
            CodeExpression condition = SerializeExpression(input.Expression);
            CodeConditionStatement output = new CodeConditionStatement(condition);

            list.Add(output);
            return list;
        }

        private CodeStatementCollection SerializeLineIfBlockStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            LineIfStatement input = (LineIfStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);
            
            CodeExpression condition = SerializeExpression(input.Expression);
            CodeStatementCollection trueStatements = SerializeStatementCollection(input.IfStatements);
            CodeStatementCollection falseStatements = SerializeStatementCollection(input.ElseStatements);

            CodeConditionStatement output = new CodeConditionStatement();
            output.Condition = condition;
            output.TrueStatements.AddRange(trueStatements);
            output.FalseStatements.AddRange(falseStatements);

            list.Add(output);
            return list;
        }

        private CodeStatementCollection SerializeElseIfBlockStatement(Tree tree)
        {
            ElseIfBlockStatement input = (ElseIfBlockStatement)tree;
            CodeStatementCollection list = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            list.AddRange(commentStatements);

            CodeStatementCollection condition = SerializeStatement(input.ElseIfStatement);
            CodeStatementCollection statements = SerializeStatementCollection(input.Statements);

            CodeConditionStatement output = null;
            foreach (CodeStatement item in condition)
            {
                if (item is CodeConditionStatement)
                {
                    output = (CodeConditionStatement)item;
                    break;
                }
            }
            if (output == null)
                output = new CodeConditionStatement(new CodePrimitiveExpression(true));

            output.TrueStatements.AddRange(statements);
            list.Add(output);
            return list;
        }

        private CodeStatementCollection SerializeEmptyStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            EmptyStatement input = (EmptyStatement)tree;
            CodeStatementCollection output = new CodeStatementCollection();

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            output = ConvertComments(comments);

            return output;            
        }

        private CodeStatementCollection SerializeEndStatement(Tree tree)
        {
            if (tree == null)
                return new CodeStatementCollection();

            EndStatement input = (EndStatement)tree;
            CodeStatementCollection output = new CodeStatementCollection();

            //input.Comments

            CodeVariableReferenceExpression variable = new CodeVariableReferenceExpression("Application");
            CodeMethodInvokeExpression invoke = new CodeMethodInvokeExpression(variable, "Exit");
            output.Add(invoke);

            return output;
        }

        #endregion

        #region Declarations

        private CodeDeclaration SerializeDeclaration(Tree tree)
        {
            if (tree == null)
                return null;

            CodeDeclaration output = null;
            
            switch (tree.Type)
            {
                case TreeType.EndBlockDeclaration:
                    output = SerializeEndBlockDeclaration(tree);
                    break;

                case TreeType.EventDeclaration:
                    output = SerializeEventDeclaration(tree);
                    break;

                case TreeType.CustomEventDeclaration:
                    output = SerializeCustomEventDeclaration(tree);
                    break;

                case TreeType.ConstructorDeclaration:
                case TreeType.SubDeclaration:
                case TreeType.FunctionDeclaration:
                    output = SerializeMethodDeclaration(tree);
                    break;

                case TreeType.OperatorDeclaration:
                    output = SerializeOperatorDeclaration(tree);
                    break;

                case TreeType.ExternalSubDeclaration:
                case TreeType.ExternalFunctionDeclaration:
                    output = SerializeExternalDeclaration(tree);
                    break;

                case TreeType.PropertyDeclaration:
                    output = SerializePropertyDeclaration(tree);
                    break;

                case TreeType.GetAccessorDeclaration:
                    output = SerializeGetAccessorDeclaration(tree);
                    break;

                case TreeType.SetAccessorDeclaration:
                    output = SerializeSetAccessorDeclaration(tree);
                    break;

                case TreeType.EnumDeclaration:
                    output = SerializeEnumDeclaration(tree);
                    break;

                case TreeType.EnumValueDeclaration:
                    output = SerializeEnumValueDeclaration(tree);
                    break;

                case TreeType.DelegateSubDeclaration:
                    output = SerializeDelegateSubDeclaration(tree);
                    break;

                case TreeType.DelegateFunctionDeclaration:
                    output = SerializeDelegateFunctionDeclaration(tree);
                    break;

                case TreeType.ModuleDeclaration:
                    output = SerializeModuleDeclaration(tree);
                    break;

                case TreeType.ClassDeclaration:
                    output = SerializeClassDeclaration(tree);
                    break;

                case TreeType.StructureDeclaration:
                    output = SerializeStructureDeclaration(tree);
                    break;

                case TreeType.InterfaceDeclaration:
                    output = SerializeInterfaceDeclaration(tree);
                    break;

                case TreeType.NamespaceDeclaration:
                    output = SerializeNamespaceDeclaration(tree);
                    break;

                case TreeType.OptionDeclaration:
                    output = SerializeOptionDeclaration(tree);
                    break;

                case TreeType.VariableListDeclaration:
                    output = SerializeVariableListDeclaration(tree);
                    break;

                case TreeType.EmptyDeclaration:
                    output = SerializeEmptyDeclaration(tree);
                    break;

                case TreeType.ImplementsDeclaration:
                    output = SerializeImplementsDeclaration(tree);
                    break;

                default:
                    throw new NotImplementedException("SerializeDeclaration not implemented for class " + tree.GetType().Name);
            }

            return output;
        }

        private CodeDeclaration SerializeEndBlockDeclaration(Tree tree)
        {
            throw new NotImplementedException();
            //EndBlockDeclaration input = (EndBlockDeclaration)tree;
            //input.Comments
            //input.EndType
            //input.EndArgumentLocation
        }

        private CodeDeclaration SerializeEventDeclaration(Tree tree)
        {
            if (tree == null)
                return null;

            EventDeclaration input = (EventDeclaration)tree;
            CodeDeclaration output = null;

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            MemberAttributes modifiers = SerializeModifier(input.Modifiers);
            CodeParameterCollection parameters = SerializeParameterCollection(input.Parameters);
            
            CodeReferenceName eventName = SerializeName(input.Name);
            string delegateName = eventName.Name + "Handler";

            CodeMemberEvent @event = new CodeMemberEvent();
            CodeTypeDelegate @delegate = new CodeTypeDelegate();

            @event.Attributes = modifiers;
            @event.Comments.AddRange(comments);
            @event.Name = eventName.Name;
            @event.Type = new CodeTypeReference(delegateName);

            @delegate.Attributes = modifiers;
            @delegate.Name = delegateName;
            @delegate.Parameters.AddRange(parameters.Parameters);

            //parameters[0].IsOptional
            //parameters.Initializers

            output = new CodeDeclaration(CodeDeclarationType.MemberCollection);
            output.Members.Add(@delegate);
            output.Members.Add(@event);
            
            //input.Attributes
            //input.ResultTypeAttributes
            //input.ResultType
            //input.ImplementsList

            return output;
        }

        private CodeDeclaration SerializeCustomEventDeclaration(Tree tree)
        {
            throw new NotImplementedException();
            //CustomEventDeclaration input = (CustomEventDeclaration)tree;
            //input.Comments
            //input.Attributes
            //input.Modifiers
            //input.CustomLocation
            //input.KeywordLocation
            //input.Name
            //input.AsLocation
            //input.ResultType
            //input.ImplementsList
            //input.Accessors
            //input.EndDeclaration
        }

        private CodeDeclaration SerializeMethodDeclaration(Tree tree)
        {
            if (tree == null)
                return null;

            MethodDeclaration input = (MethodDeclaration)tree;
            CodeDeclaration output = new CodeDeclaration(CodeDeclarationType.MemberCollection);

            //input.Attributes
            //input.ImplementsList
            //input.HandlesList
            //input.ResultTypeAttributes
            //input.TypeParameters

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            MemberAttributes modifiers = SerializeModifierCollection(input.Modifiers);
            CodeReferenceName reference = SerializeSimpleName(input.Name);
            CodeParameterCollection parameters = SerializeParameterCollection(input.Parameters);
            CodeStatementCollection statements = SerializeStatementCollection(input.Statements);
            CodeTypeReference returnType = null;
            
            CodeMemberMethod member = null;

            switch (tree.Type)
            {
                case TreeType.SubDeclaration:
                    member = new CodeMemberMethod();

                    if (string.Compare(reference.Name, "Main", true) == 0)
                        modifiers = modifiers | MemberAttributes.Static;

                    break;

                case TreeType.FunctionDeclaration:
                    member = new CodeMemberMethod();
                    returnType = SerializeTypeName(input.ResultType);
                    break;

                case TreeType.ConstructorDeclaration:
                    member = new CodeConstructor();
                    break;
            }
            
            if (returnType != null)
                AddReturnStatements(returnType, reference.Name, statements);
            

            modifiers = modifiers | MemberAttributes.Final;
                        
            member.Attributes = modifiers;
            member.Comments.AddRange(comments);
            member.Name = reference.Name;
            member.Parameters.AddRange(parameters.Parameters);
            member.ReturnType = returnType;
            member.Statements.AddRange(statements);

            output.Members.Add(member);
            return output;
        }

        private CodeDeclaration SerializeOperatorDeclaration(Tree tree)
        {
            throw new NotImplementedException();
            //OperatorDeclaration input = (OperatorDeclaration)tree;
            //input.Comments
            //input.Attributes
            //input.Modifiers
            //input.KeywordLocation
            //input.OperatorToken
            //input.Parameters
            //input.AsLocation
            //input.ResultTypeAttributes
            //input.ResultType
            //input.Statements
            //input.EndDeclaration
        }

        
        private CodeDeclaration SerializeExternalDeclaration(Tree tree)
        {
            if (tree == null)
                return null;

            ExternalDeclaration input = (ExternalDeclaration)tree;
            CodeDeclaration output = new CodeDeclaration(CodeDeclarationType.MemberCollection);

            //[DllImport("User32.dll", EntryPoint="Test")]
            //public static extern void Test();

            //input.Attributes
            //input.Charset
            //input.ResultTypeAttributes

            string libLocation = string.Empty;
            string aliasName = string.Empty;

            if (input.LibLiteral != null)
                libLocation = input.LibLiteral.Literal;

            if (input.AliasLiteral != null)
                aliasName = input.AliasLiteral.Literal;
            
            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            MemberAttributes modifiers = SerializeModifierCollection(input.Modifiers);
            CodeReferenceName reference = SerializeSimpleName(input.Name);
            CodeParameterCollection parameters = SerializeParameterCollection(input.Parameters);
            
            CodeTypeReference returnType = null;
            CodeMemberMethod member = null;

            switch (tree.Type)
            {
                case TreeType.ExternalSubDeclaration:
                    member = new CodeMemberMethod();
                    break;

                case TreeType.ExternalFunctionDeclaration:
                    member = new CodeMemberMethod();
                    returnType = SerializeTypeName(input.ResultType);
                    break;
            }

            CodeTypeReference type = new CodeTypeReference("DllImport");
            CodeAttributeDeclaration attribute = new CodeAttributeDeclaration(type);

            if (!string.IsNullOrEmpty(libLocation))
            {
                CodeAttributeArgument dllLocation = new CodeAttributeArgument(new CodePrimitiveExpression(libLocation));
                attribute.Arguments.Add(dllLocation);
            }

            if (!string.IsNullOrEmpty(aliasName))
            {
                CodeAttributeArgument entryPoint = new CodeAttributeArgument("EntryPoint", new CodePrimitiveExpression(aliasName));
                attribute.Arguments.Add(entryPoint);
            }

            member.CustomAttributes.Add(attribute);
            member.Attributes = modifiers;
            member.Comments.AddRange(comments);
            member.Name = reference.Name;
            member.Parameters.AddRange(parameters.Parameters);
            member.ReturnType = returnType;
            
            output.Members.Add(member);
            return output;
        }

        private CodeDeclaration SerializePropertyDeclaration(Tree tree)
        {
            if (tree == null)
                return null;

            PropertyDeclaration input = (PropertyDeclaration)tree;
            CodeDeclaration output = new CodeDeclaration(CodeDeclarationType.MemberCollection);

            //input.Attributes
            //input.ImplementsList
            //input.ResultTypeAttributes

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            MemberAttributes modifiers = SerializeModifierCollection(input.Modifiers);
            CodeReferenceName reference = SerializeSimpleName(input.Name);
            CodeParameterCollection parameters = SerializeParameterCollection(input.Parameters);
            CodeTypeReference returnType = SerializeTypeName(input.ResultType);

            CodeDeclaration getAccessor = SerializeGetAccessorDeclaration(input.GetAccessor);
            CodeDeclaration setAccessor = SerializeSetAccessorDeclaration(input.SetAccessor);

            modifiers = modifiers | MemberAttributes.Final;

            CodeMemberProperty member = new CodeMemberProperty();
            member.Comments.AddRange(comments);
            member.Attributes = modifiers;
            member.Name = reference.Name;
            member.Parameters.AddRange(parameters.Parameters);

            if (getAccessor != null)
                member.GetStatements.AddRange(getAccessor.Statements);
            if (setAccessor != null)
                member.SetStatements.AddRange(setAccessor.Statements);
            member.Type = returnType;
            
            output.Members.Add(member);
            return output;
        }

        private CodeDeclaration SerializeGetAccessorDeclaration(Tree tree)
        {
            if (tree == null)
                return null;

            GetAccessorDeclaration input = (GetAccessorDeclaration)tree;
            CodeDeclaration output = new CodeDeclaration(CodeDeclarationType.StatementCollection);

            //input.Attributes
            //input.Modifiers

            CodeTypeReference returnType = new CodeTypeReference(typeof(object));
            string propertyName = string.Empty;

            if (input.Parent is PropertyDeclaration)
            {
                PropertyDeclaration property = (PropertyDeclaration)input.Parent;
                returnType = SerializeTypeName(property.ResultType);
                
                if (property.Name != null)
                    propertyName = property.Name.Name;
            }

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            CodeStatementCollection statements = SerializeStatementCollection(input.Statements);

            AddReturnStatements(returnType, propertyName, statements);

            output.Statements.AddRange(commentStatements);
            output.Statements.AddRange(statements);
           
            return output;
        }

        private CodeDeclaration SerializeSetAccessorDeclaration(Tree tree)
        {
            if (tree == null)
                return null;

            SetAccessorDeclaration input = (SetAccessorDeclaration)tree;
            CodeDeclaration output = new CodeDeclaration(CodeDeclarationType.StatementCollection);

            //input.Attributes
            //input.Modifiers

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            CodeStatementCollection commentStatements = ConvertComments(comments);
            CodeStatementCollection statements = SerializeStatementCollection(input.Statements);

            CodeParameterCollection parameters = SerializeParameterCollection(input.Parameters);
            CodeParameterDeclarationExpression setParameter = null;
            if (parameters != null && parameters.Count > 0)
                setParameter = parameters.Parameters[0];

            foreach (CodeStatement statement in statements)
            {
                CodeAssignStatement assign = statement as CodeAssignStatement;

                if (assign != null)
                {
                    CodeVariableReferenceExpression variable = assign.Right as CodeVariableReferenceExpression;
                    if (variable != null && string.Compare(variable.VariableName, setParameter.Name, true) == 0)
                    {
                        assign.Right = new CodePropertySetValueReferenceExpression();
                        break;
                    }
                }
            }

            output.Statements.AddRange(commentStatements);
            output.Statements.AddRange(statements);

            return output;
        }

        private CodeDeclaration SerializeEnumDeclaration(Tree tree)
        {
            if (tree == null)
                return null;

            EnumDeclaration input = (EnumDeclaration)tree;
            CodeDeclaration output = new CodeDeclaration(CodeDeclarationType.MemberCollection);

            //input.Attributes
            //input.ElementType

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            MemberAttributes modifiers = SerializeModifierCollection(input.Modifiers);
            CodeReferenceName reference = SerializeSimpleName(input.Name);
            CodeDeclarationCollection declarations = SerializeDeclarationCollection(input.Declarations);
            
            CodeTypeDeclaration declaration = new CodeTypeDeclaration();
            declaration.IsEnum = true;
            declaration.Comments.AddRange(comments);
            declaration.Attributes = modifiers;
            declaration.Name = reference.Name;
            declaration.Members.AddRange(declarations.Members);

            output.Members.Add(declaration);
            return output;
        }

        private CodeDeclaration SerializeEnumValueDeclaration(Tree tree)
        {
            if (tree == null)
                return null;

            EnumValueDeclaration input = (EnumValueDeclaration)tree;
            CodeDeclaration output = new CodeDeclaration(CodeDeclarationType.MemberCollection);

            //input.Attributes

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            MemberAttributes modifiers = SerializeModifierCollection(input.Modifiers);
            CodeReferenceName reference = SerializeSimpleName(input.Name);
            CodeExpression expression = null;

            if (input.Expression != null)
                expression = SerializeExpression(input.Expression);

            CodeMemberField member = new CodeMemberField();
            
            member.Comments.AddRange(comments);
            member.Attributes = modifiers;
            member.Name = reference.Name;
            if (expression != null)
                member.InitExpression = expression;

            output.Members.Add(member);
            return output;            
        }

        private CodeDeclaration SerializeDelegateSubDeclaration(Tree tree)
        {
            throw new NotImplementedException();
            //DelegateSubDeclaration input = (DelegateSubDeclaration)tree;
            //input.Comments
            //input.Attributes
            //input.Modifiers
            //input.KeywordLocation
            //input.SubOrFunctionLocation
            //input.Name
            //input.TypeParameters
            //input.Parameters
        }

        private CodeDeclaration SerializeDelegateFunctionDeclaration(Tree tree)
        {
            throw new NotImplementedException();
            //DelegateFunctionDeclaration input =(DelegateFunctionDeclaration)tree;
            //input.Comments
            //input.Attributes
            //input.Modifiers
            //input.Name
            //input.Parameters
            //input.ResultTypeAttributes
            //input.ResultType
        }

        private CodeDeclaration SerializeModuleDeclaration(Tree tree)
        {
            throw new NotImplementedException();
            //ModuleDeclaration input = (ModuleDeclaration)tree;
            //input.Comments
            //input.Attributes
            //input.Modifiers
            //input.Name
            //input.Declarations
        }

        private CodeDeclaration SerializeClassDeclaration(Tree tree)
        {
            throw new NotImplementedException();
            //ClassDeclaration input = (ClassDeclaration)tree;
            //input.Comments
            //input.Attributes
            //input.Modifiers
            //input.Name
            //input.TypeParameters
            //input.Declarations
        }

        private CodeDeclaration SerializeStructureDeclaration(Tree tree)
        {
            if (tree == null)
                return null;

            StructureDeclaration input = (StructureDeclaration)tree;
            CodeDeclaration output = new CodeDeclaration(CodeDeclarationType.MemberCollection);

            //input.Attributes
            //input.TypeParameters

            CodeTypeDeclaration member = new CodeTypeDeclaration();
            member.IsStruct = true;

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            MemberAttributes modifiers = SerializeModifierCollection(input.Modifiers);
            CodeReferenceName reference = SerializeSimpleName(input.Name);
            CodeDeclarationCollection declarations = SerializeDeclarationCollection(input.Declarations);

            member.Comments.AddRange(comments);
            member.Attributes = modifiers;
            member.Name = reference.Name;
            member.Members.AddRange(declarations.Members);

            output.Members.Add(member);
            return output;
        }

        private CodeDeclaration SerializeInterfaceDeclaration(Tree tree)
        {
            throw new NotImplementedException();
            //InterfaceDeclaration input = (InterfaceDeclaration)tree;
            //input.Comments
            //input.Attributes
            //input.Modifiers
            //input.Name
            //input.TypeParameters
            //input.Declarations
        }

        private CodeDeclaration SerializeNamespaceDeclaration(Tree tree)
        {
            throw new NotImplementedException();
            //NamespaceDeclaration input = (NamespaceDeclaration)tree;
            //input.Comments
            //input.Attributes
            //input.Modifiers
            //input.NamespaceLocation
            //input.Name
            //input.Declarations
            //input.EndDeclaration
        }

        private CodeDeclaration SerializeOptionDeclaration(Tree tree)
        {
            if (tree == null)
                return null;

            OptionDeclaration input = (OptionDeclaration)tree;
            CodeDeclaration output = new CodeDeclaration(CodeDeclarationType.StatementCollection);

            CodeCommentStatement comment = null;

            switch (input.OptionType)
            {
                case OptionType.Explicit:
                case OptionType.ExplicitOn:
                case OptionType.ExplicitOff:
                    comment = new CodeCommentStatement("TODO: Option Explicit Statement is not supported");
                    break;
                case OptionType.Strict:
                case OptionType.StrictOn:
                case OptionType.StrictOff:
                    throw new NotImplementedException();
                case OptionType.CompareBinary:
                case OptionType.CompareText:
                    comment = new CodeCommentStatement("INFO: Option Compare: " + Enum.GetName(typeof(OptionType), input.OptionType));
                    VisualBasicOptions.Add(input.OptionType);
                    break;
                case OptionType.BaseZero:
                case OptionType.BaseOne:
                    comment = new CodeCommentStatement("INFO: Option Base: " + Enum.GetName(typeof(OptionType), input.OptionType));
                    VisualBasicOptions.Add(input.OptionType);
                    break;
            }

            if (comment != null)
                output.Statements.Add(comment);

            return output;
        }

        private CodeDeclaration SerializeVariableListDeclaration(Tree tree)
        {
            if (tree == null)
                return null;

            VariableListDeclaration input = (VariableListDeclaration)tree;
            CodeDeclaration output = null;

            CodeTypeMemberCollection members = null;

            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            MemberAttributes modifiers = SerializeModifierCollection(input.Modifiers);
            CodeVariableDeclaratorCollection declarations = SerializeVariableDeclaratorCollection(input.VariableDeclarators);

            if (declarations.CollectionType == CodeCollectionType.TypeMember)
                members = declarations.Members;
            else
                throw new Exception("Invalid collection in SerializeVariableListDeclaration: " + declarations.GetType().Name);

            int position = 0;
            foreach (CodeTypeMember field in members)
            {
                if (++position == 0)
                    field.Comments.AddRange(comments);

                field.Attributes = modifiers;
            }

            output = new CodeDeclaration(members);

            return output;
        }

        private CodeDeclaration SerializeEmptyDeclaration(Tree tree)
        {
            if (tree == null)
                return null;

            EmptyDeclaration input = (EmptyDeclaration)tree;
            CodeDeclaration output = new CodeDeclaration(CodeDeclarationType.CommentCollection);
            
            CodeCommentStatementCollection comments = SerializeComments(input.Comments);
            output.Comments.AddRange(comments);

            return output;
        }

        private CodeDeclaration SerializeImplementsDeclaration(Tree tree)
        {
            //ImplementsDeclaration input = (ImplementsDeclaration)tree;
            CodeDeclaration output = new CodeDeclaration(CodeDeclarationType.StatementCollection);
            CodeCommentStatement comment = new CodeCommentStatement("TODO: ImplementsDeclaration is not supported");
            output.Comments.Add(comment);
            return output;
        }

        #endregion

        #region Public Methods

        public CodeTypeDeclaration Convert(Tree tree)
        {
            CodeTypeDeclaration output = null;

            if (Tree.Types.IsFile(tree))
                output = ConvertFile(tree);
            else if (Tree.Types.IsDeclaration(tree))
                output = ConvertDeclaration(tree);
            else if (Tree.Types.IsStatement(tree))
                output = ConvertStatement(tree);
            else if (Tree.Types.IsExpression(tree))
                output = ConvertExpression(tree);
            else
                throw new NotImplementedException("Convert not implemented for the tree " + tree.Type.ToString());

            return output;
        }

        private CodeTypeDeclaration ConvertFile(Tree tree)
        {
            if (tree == null)
                return null;

            FileTree input = (FileTree)tree;

            CodeTypeDeclaration output = new CodeTypeDeclaration();
            output.IsClass = true;
            output.Name = string.IsNullOrEmpty(input.Name) ? "FileResult" : input.Name;;

            CodeDeclarationCollection declarations = SerializeDeclarationCollection(input.Declarations);

            output.Comments.AddRange(declarations.Comments);
            output.Members.AddRange(declarations.Members);
            
            return output;
        }

        private CodeTypeDeclaration ConvertDeclaration(Tree tree)
        {
            if (tree == null)
                return null;

            Declaration input = (Declaration)tree;

            CodeTypeDeclaration output = new CodeTypeDeclaration();
            output.IsClass = true;
            output.Name = "DeclarationResult";

            CodeDeclaration declaration = SerializeDeclaration(input);

            output.Comments.AddRange(declaration.Comments);
            output.Members.AddRange(declaration.Members);

            return output;
        }

        private CodeTypeDeclaration ConvertStatement(Tree tree)
        {
            if (tree == null)
                return null;

            Statement input = (Statement)tree;

            CodeTypeDeclaration output = new CodeTypeDeclaration();
            output.IsClass = true;
            output.Name = "StatementResult";

            CodeStatementCollection statements = SerializeStatement(input);
            
            CodeMemberMethod member = new CodeMemberMethod();
            member.Name = "TestMethod";
            member.Statements.AddRange(statements);

            output.Members.Add(member);

            return output;
        }

        private CodeTypeDeclaration ConvertExpression(Tree tree)
        {
            if (tree == null)
                return null;

            Expression input = (Expression)tree;

            CodeTypeDeclaration output = new CodeTypeDeclaration();
            output.IsClass = true;
            output.Name = "ExpressionResult";

            CodeExpression expression = SerializeExpression(input);

            CodeMemberMethod member = new CodeMemberMethod();
            member.Name = "TestMethod";
            member.Statements.Add(expression);

            output.Members.Add(member);

            return output;
        }

        #endregion

        #region Code Generation

        public CodeCompileUnit CreateCompileUnit(CodeNamespace ns)
        {
            if (ns == null)
                throw new ArgumentNullException("The namespace must not be null");

            CodeCompileUnit unit = new CodeCompileUnit();
            unit.Namespaces.Add(ns);

            CompilerParameters parameters = new CompilerParameters();

            parameters.ReferencedAssemblies.Add("mscorlib.dll");
            parameters.ReferencedAssemblies.Add("System.dll");
            parameters.ReferencedAssemblies.Add("System.Drawing.dll");
            parameters.ReferencedAssemblies.Add("System.Windows.Forms.dll");

            return unit;
        }

        public CodeNamespace CreateNamespace(string name, CodeTypeDeclaration classTree)
        {
            if (classTree == null)
                throw new ArgumentNullException("The class tree must not be null");

            CodeNamespace ns = new CodeNamespace(name);
            ns.Imports.Add(new CodeNamespaceImport("System"));
            ns.Imports.Add(new CodeNamespaceImport("System.Collections"));
            ns.Imports.Add(new CodeNamespaceImport("System.Collections.Generic"));
            ns.Imports.Add(new CodeNamespaceImport("System.Text"));
            ns.Imports.Add(new CodeNamespaceImport("System.Runtime.InteropServices"));
            ns.Imports.Add(new CodeNamespaceImport("System.Windows.Forms"));
            ns.Imports.Add(new CodeNamespaceImport("Microsoft.VisualBasic"));
            ns.Imports.Add(new CodeNamespaceImport("Microsoft.VisualBasic.Compatibility.VB6"));
            ns.Imports.Add(new CodeNamespaceImport("Microsoft.VisualBasic.Devices"));
            ns.Imports.Add(new CodeNamespaceImport("Microsoft.VisualBasic.FileIO"));
            ns.Imports.Add(new CodeNamespaceImport("Microsoft.VisualBasic.Logging"));
            
            ns.Types.Add(classTree);

            return ns;
        }

        #endregion
    }
}

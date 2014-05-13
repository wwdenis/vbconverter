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
namespace VBConverter.CodeParser.Trees
{
	/// <summary>
	/// The type of a tree.
	/// </summary>
	public enum TreeType
	{
		SyntaxError,

		// Collections
		ArgumentCollection,
		ExpressionCollection,
		InitializerCollection,
		VariableNameCollection,
		VariableDeclaratorCollection,
		ParameterCollection,
		TypeParameterCollection,
		TypeArgumentCollection,
		TypeConstraintCollection,
		CaseClauseCollection,
		AttributeCollection,
		AttributeBlockCollection,
		NameCollection,
		TypeNameCollection,
		ImportCollection,
		ModifierCollection,
		StatementCollection,
		DeclarationCollection,

		// Comments
		Comment,

		// Names
		SimpleName,
		VariableName,
		QualifiedName,
		GlobalNamespaceName,
		MeName,
		MyBaseName,

		// Types
		IntrinsicType,
		NamedType,
		ConstructedType,
		ArrayType,

		// Initializers
		AggregateInitializer,
		ExpressionInitializer,

		// Arguments
		Argument,

		// Expressions
		SimpleNameExpression,
		TypeReferenceExpression,
		QualifiedExpression,
		DictionaryLookupExpression,
		GenericQualifiedExpression,
		CallOrIndexExpression,
		NewExpression,
		NewAggregateExpression,
		StringLiteralExpression,
		CharacterLiteralExpression,
		DateLiteralExpression,
		IntegerLiteralExpression,
		FloatingPointLiteralExpression,
		DecimalLiteralExpression,
		BooleanLiteralExpression,
		BinaryOperatorExpression,
		UnaryOperatorExpression,
		AddressOfExpression,
		IntrinsicCastExpression,
		InstanceExpression,
		GlobalExpression,
		NothingExpression,
		ParentheticalExpression,
		CTypeExpression,
		DirectCastExpression,
		TryCastExpression,
		TypeOfExpression,
		GetTypeExpression,

		// Statements
		EmptyStatement,
		GotoStatement,
        GoSubStatement,
		ExitStatement,
		ContinueStatement,
		StopStatement,
		EndStatement,
		ReturnStatement,
		RaiseEventStatement,
		AddHandlerStatement,
		RemoveHandlerStatement,
		ErrorStatement,
		OnErrorStatement,
		ResumeStatement,
		ReDimStatement,
		EraseStatement,
		CallStatement,
		AssignmentStatement,
		CompoundAssignmentStatement,
		MidAssignmentStatement,
		LocalDeclarationStatement,
		LabelStatement,
		DoBlockStatement,
		ForBlockStatement,
		ForEachBlockStatement,
		WhileBlockStatement,
		SyncLockBlockStatement,
		UsingBlockStatement,
		WithBlockStatement,
		IfBlockStatement,
		ElseIfBlockStatement,
		ElseBlockStatement,
		LineIfBlockStatement,
		ThrowStatement,
		TryBlockStatement,
		CatchBlockStatement,
		FinallyBlockStatement,
		SelectBlockStatement,
		CaseBlockStatement,
		CaseElseBlockStatement,
		LoopStatement,
		NextStatement,
		CatchStatement,
		FinallyStatement,
		CaseStatement,
		CaseElseStatement,
		ElseIfStatement,
		ElseStatement,
		EndBlockStatement,

		// Modifiers
		Modifier,

		// VariableDeclarators
		VariableDeclarator,

		// CaseClauses
		ComparisonCaseClause,
		RangeCaseClause,

		// Attributes
		AttributeTree,

		// Declaration statements
		EmptyDeclaration,
		NamespaceDeclaration,
		VariableListDeclaration,
		SubDeclaration,
		FunctionDeclaration,
		OperatorDeclaration,
		ConstructorDeclaration,
		ExternalSubDeclaration,
		ExternalFunctionDeclaration,
		PropertyDeclaration,
		GetAccessorDeclaration,
		SetAccessorDeclaration,
		EventDeclaration,
		CustomEventDeclaration,
		AddHandlerAccessorDeclaration,
		RemoveHandlerAccessorDeclaration,
		RaiseEventAccessorDeclaration,
		EndBlockDeclaration,
		InheritsDeclaration,
		ImplementsDeclaration,
		ImportsDeclaration,
		OptionDeclaration,
		AttributeDeclaration,
		EnumValueDeclaration,
		ClassDeclaration,
		StructureDeclaration,
		ModuleDeclaration,
		InterfaceDeclaration,
		EnumDeclaration,
		DelegateSubDeclaration,
		DelegateFunctionDeclaration,

		// Parameters
		Parameter,

		// Imports
		NameImport,
		AliasImport,

		// Type Parameters
		TypeParameter,

		// Files
		FileTree
	}
}

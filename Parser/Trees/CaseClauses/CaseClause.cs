using System;
using System.Collections.Generic;
using System.Diagnostics;
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

using VBConverter.CodeParser.Trees.Expressions;
using VBConverter.CodeParser.Trees.Statements;
using VBConverter.CodeParser.Trees.Types;

namespace VBConverter.CodeParser.Trees.CaseClause
{
	/// <summary>
	/// A parse tree for a case clause in a Select statement.
	/// </summary>
	public abstract class CaseClause : Tree
	{
        public abstract BinaryOperatorExpression ComparisonExpression
        {
            get;
        }

        public SelectBlockStatement SelectCaseBlock
        {
            get
            {
                if (Parent == null ||
                    !(Parent is CaseClauseCollection) ||
                    Parent.Parent == null ||
                    !(Parent.Parent is CaseStatement) ||
                    Parent.Parent.Parent == null ||
                    !(Parent.Parent.Parent is CaseBlockStatement) ||
                    Parent.Parent.Parent.Parent == null ||
                    !(Parent.Parent.Parent.Parent is StatementCollection) ||
                    Parent.Parent.Parent.Parent.Parent == null ||
                    !(Parent.Parent.Parent.Parent.Parent is SelectBlockStatement)
                    )
                    return null;

                return (SelectBlockStatement)Parent.Parent.Parent.Parent.Parent;
            }
        }

        protected CaseClause(TreeType type, Span span) : base(type, span)
		{
			Debug.Assert(type == TreeType.ComparisonCaseClause || type == TreeType.RangeCaseClause);
		}
	}
}

using System;
using System.Collections.Generic;
using System.Text;
namespace VBConverter.CodeParser.Trees.Types
{
	public enum PrecedenceLevel
	{
		None,
		Xor,
		Or,
		And,
		Not,
		Relational,
		Shift,
		Concatenate,
		Plus,
		Modulus,
		IntegralDivide,
		Multiply,
		Negate,
		Power,
		Range
	}
}

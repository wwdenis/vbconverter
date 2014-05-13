using System;
using System.Collections.Generic;
using System.Text;

namespace VBConverter.CodeParser.Context
{
	internal sealed class RegionContext
	{
		private Location _Start;
		private string _Description;

		public Location Start {
			get { return _Start; }
			set { _Start = value; }
		}

		public string Description {
			get { return _Description; }
			set { _Description = value; }
		}
	}
}

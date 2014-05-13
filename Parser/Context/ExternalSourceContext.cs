using System;
using System.Collections.Generic;
using System.Text;

namespace VBConverter.CodeParser.Context
{
	internal sealed class ExternalSourceContext
	{
		private Location _Start;
		private string _File;
		private long _Line;

		public Location Start {
			get { return _Start; }
			set { _Start = value; }
		}

		public string File {
			get { return _File; }
			set { _File = value; }
		}

		public long Line {
			get { return _Line; }
			set { _Line = value; }
		}
	}
}

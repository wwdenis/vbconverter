using System;
using System.Collections.Generic;
using System.Text;

namespace VBConverter.CodeParser.Context
{
	internal sealed class ConditionalCompilationContext
	{
		private bool _BlockActive;
		private bool _AnyBlocksActive;
		private bool _SeenElse;

		public bool BlockActive {
			get { return _BlockActive; }
			set { _BlockActive = value; }
		}

		public bool AnyBlocksActive {
			get { return _AnyBlocksActive; }
			set { _AnyBlocksActive = value; }
		}

		public bool SeenElse {
			get { return _SeenElse; }
			set { _SeenElse = value; }
		}
	}
}

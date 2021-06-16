using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Resources;
using System.Reflection;
using System.Globalization;
using Word = Microsoft.Office.Interop.Word;
using resCommon = GeneratorV2.Properties.CommonData;
using resUser = GeneratorV2.Properties.UserData;
using resDocument = GeneratorV2.Properties.DocumentData;

[assembly:NeutralResourcesLanguage("ru-RU")]

namespace GeneratorV2
{
	class Generator
	{
		static readonly object oMissing = Missing.Value;
		static readonly object oEndOfDoc = "\\endofdoc";

		static Word.Application app;
		Word.Document doc;

		static void Main(string[] args)
		{
			try
			{
				app = new Word.Application();
				app.Visible = true;
			}
			catch(Exception e)
            {

            }
		}
	}
}

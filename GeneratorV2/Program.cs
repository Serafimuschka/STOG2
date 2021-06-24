using System;
using System.Linq;
using System.Resources;
using System.Reflection;
using Microsoft.Win32;
using Word = Microsoft.Office.Interop.Word;

using resCommon = GeneratorV2.Properties.CommonData;
using resUser = GeneratorV2.Properties.UserData;
using resDocument = GeneratorV2.Properties.DocumentData;

[assembly:NeutralResourcesLanguage("ru-RU")]

namespace GeneratorV2
{
	public struct ConstantIndents
    {
		public const float __1p00 = 28.3465F;
		public const float __1p25 = 35.43307F;
		public const float __2p50 = 70.8661F;
	}
	public struct StyleSet
    {
		public static readonly string __sto_sh1 = "60-02.2.3-2018 Section H1";
		public static readonly string __sto_sh2 = "60-02.2.3-2018 Section H2";
		public static readonly string __sto_dh1 = "60-02.2.3-2018 Headers D1";
		public static readonly string __sto_dft = "60-02.2.3-2018 Para DFT";
	}
	class Generator
	{
		static object oMissing = Missing.Value;
		static object oEndOfDoc = "\\endofdoc";
		static object g_style = "GlobalStyle";

		static string[] acceptedArgs = 
		{
			"--title",
			"--debug"
		};

		static Word.Application app;
		static Word.Document doc;

		static string uSurname;
		static string uForename;
		static string uPatronymic;
		static string uGender;
		static string uCourse;
		static string uGroup;
		static string uCode;
		static string uDirection;
		static string uHighSchool;

		static string dYear;
		static string dDiscipline;
		static string dTheme;
		static string dPrepod;
		static string dPrepodIniz;
		static string dPrepodInfo;

		static void Main(string[] args)
		{
			Console.Title = "STOGv2 ver. 2.0.0.0";

			try
			{
				Console.ForegroundColor = ConsoleColor.Green;
				Console.WriteLine("STOGv2 Generator module v.2.0");
				Console.WriteLine("Initialization: reading args...");

				if (args.Length == 0)
				{
					Console.ForegroundColor = ConsoleColor.Red;
					Console.WriteLine("> Invalid args.");
					Console.ResetColor();
					Help();
				}
				else if (args.Length > 0)
				{
					var cmd = args[0];
					switch (cmd)
					{
						default:
							Console.WriteLine("> Unknown args.");
							Console.Write("Accepted: ");
							
							for (int i = 0; i < args.Length; i++)
							{
								if (!acceptedArgs.Contains(args[i]))
								{
									Console.ForegroundColor = ConsoleColor.Red;
									Console.BackgroundColor = ConsoleColor.Gray;
								}	
								Console.Write(args[i] + ' ');
								Console.ForegroundColor = ConsoleColor.Green;
								Console.BackgroundColor = ConsoleColor.Black;
							}

							Console.WriteLine();
							Help();
							break;
						case "--title" when args.Length >= 3:
							int pattern = Convert.ToInt32(args[1]);
							int model = Convert.ToInt32(args[2]);

							Console.WriteLine
							(
								"\nAccepted command <--title> with params:"
							);
							Console.WriteLine
							(
								"\t PAT TYPE = " + '\t' + pattern.ToString()
							);
							Console.WriteLine
							(
								"\t MDL TYPE = " + '\t' + model.ToString()
							);
							Console.WriteLine();

							Title(pattern, model);
							break;
					}

					app.Visible = true;
				}
			}
			catch(Exception e)
			{
				WriteException(e);
				if (!Array.Exists(args, s => s.Equals("--debug")))
				{
					Console.ReadKey();
				}
			}
			Console.ResetColor();

			if (Array.Exists(args, s => s.Equals("--debug")))
			{
				Console.WriteLine("\nCompleted. Press any button to exit.");
				Console.ReadKey();
			}
			else Environment.Exit(0);
		}

		// Reads values from the registry.
		// Returns nothing. Unlimited use.
		static void GetRegistryKeys()
		{
			Console.ForegroundColor = ConsoleColor.Yellow;
			Console.WriteLine
			(
				"::rgReader > Attempting to read registry keys..."
			);

			RegistryKey hkcu = Registry.CurrentUser;
			RegistryKey stog = hkcu.OpenSubKey("STOGv2");

			uSurname = stog.GetValue("rgUserSurname").ToString();
			uForename = stog.GetValue("rgUserForename").ToString();
			uPatronymic = stog.GetValue("rgUserPatronymic").ToString();
			uGender = stog.GetValue("rgUserGender").ToString();
			uCourse = stog.GetValue("rgUserCourse").ToString();
			uGroup = stog.GetValue("rgUserGroup").ToString();
			uCode = stog.GetValue("rgUserDirectionCode").ToString();
			uDirection = stog.GetValue("rgUserDirectionName").ToString();
			uHighSchool = stog.GetValue("rgUserHighSchool").ToString();

			dYear = stog.GetValue("rgDocYear").ToString();
			dDiscipline = stog.GetValue("rgDocLastDisc").ToString();
			dTheme = stog.GetValue("rgDocLastTheme").ToString();
			dPrepod = stog.GetValue("rgDocLastPrepod").ToString();
			dPrepodIniz = stog.GetValue("rgDocLastPrepodIniz").ToString();
			dPrepodInfo = stog.GetValue("rgDocLastPrepodInfo").ToString();

			Console.WriteLine("::rgReader > Succeeded.");
			Console.ForegroundColor = ConsoleColor.Green;

			stog.Close();
			hkcu.Close();
		}

		// Shows help message.
		static void Help()
		{
			Console.WriteLine("There are a list of generator args:");
			Console.WriteLine("\t--title p: Generates title list with type p;");
			Console.WriteLine("\t--task: Generates task list from resx;");
			Console.WriteLine("\t--note: Generates note list;");
			Console.WriteLine("\t--content: Generates contents from resx;");
			Console.WriteLine("\t--app: Generates applications from resx;");
			Console.WriteLine("\t--full: Generates full document;");
			Console.WriteLine("\t--custom bbbbb vvvvv: Custom generation.");
		}

		// Shows exception message.
		static void WriteException(Exception e)
		{
			Console.ForegroundColor = ConsoleColor.Red;
			Console.WriteLine("\n--------------------");
			Console.WriteLine("Runtime error: exception catched.");
			Console.WriteLine("Info:");
			Console.WriteLine(e.Message);
			Console.WriteLine("Task aborted.\n");
			Console.ResetColor();
		}

		// Generates title from internal pattern.
		// pattern	  :: work pattern (essay, report, etc.)
		// model	  :: pattern model (__sto_type)
		// Returns nothing. Delegates task to other funcs.
		static void Title(int pattern, int model)
		{
			/*
			 * Pattern cases:
			 * 
			 * Group Zero (01, 02, 03, 04):
			 * > 01		Essay
			 * > 02		Abstract
			 * > 03		Control work
			 * > 04		Graphic work
			 * 
			 * Group First (11, 12, 13, 14):
			 * > 11		Course proj ind.
			 * > 12		Course proj group	| Coming soon
			 * > 13		Course work ind.	| 
			 * > 14		Course work group	|
			 * 
			 * Group Second (21, 22, 23):
			 * > 21		Practics report 03
			 * > 22		Practics report 02	| Coming soon
			 * > 23		Science report		|
			 * 
			 * Group Third (31, 32, 33):
			 * > 31		Dissertation		| Coming soon
			 * > 32		Science work		|
			 * > 33		Science report		|
			 * 
			 * Group Fourth (41+):
			 *								| Coming soon
			 * 
			 */

			string __sto_work;

			Console.WriteLine
			(
				"::title > task accepted."
			);

			switch (pattern)
			{
				case 1:
					__sto_work = "Эссе";
					MakeTitleGroupOne(__sto_work, model);
					break;
				case 2:
					__sto_work = "Реферат";
					MakeTitleGroupOne(__sto_work, model);
					break;
				case 3:
					__sto_work = "Контрольная работа";
					MakeTitleGroupOne(__sto_work, model);
					break;
				case 4:
					__sto_work = "Расчётно-графическая работа";
					MakeTitleGroupOne(__sto_work, model);
					break;
				case 11:
					__sto_work = "Курсовой проект";
					MakeTitleGroupTwo(__sto_work, model);
					break;
				default:
					throw new InvalidOperationException();
			}
		}

		// Makes title page from group 0n.
		// __sto_work :: type of work (essay, abstract, etc.)
		// __sto_type :: type of discipline (discipline, module, etc.)
		// Returns nothing. Launching Word application.
		static void MakeTitleGroupOne(string __sto_work, int __sto_type = 0)
		{
			Console.WriteLine
			(
				"::title.mktg_o > working"
			);

			GetRegistryKeys();

			// Preparing block:
			string __sto_font = resDocument.__sto_font;
			string __sto_null = " ";
			int __sto_size = Convert.ToInt32(resDocument.__sto_size);

			string __sto_discA = "";
			string __sto_discB = "";
			string __sto_discC = "";

			string __sto_wtyp;
			float __sto_widA;
			float __sto_widB;

			string __sto_exec;

			var gender = Convert.ToBoolean(uGender);
			if (gender) __sto_exec = resCommon.__sto_execM;
			else __sto_exec = resCommon.__sto_execF;

			var disc = dDiscipline;
			var len = disc.Length;
			switch (__sto_type)
			{
				case 0:
					__sto_wtyp = resCommon.__sto_discA;
					__sto_widA = 99.2126F;
					__sto_widB = 393.44882F;

					if (len > 64)
					{
						__sto_discA = disc.Substring(0, 64);
						if (len > 150)
						{
							__sto_discB = disc.Substring(64, 86);
							__sto_discC = disc.Substring(150);
						}
						else __sto_discB = disc.Substring(64);
					}
					else __sto_discA = disc;

					break;
				case 1:
					__sto_wtyp = resCommon.__sto_discB;
					__sto_widA = 195.591F;
					__sto_widB = 297.07087F;

					if (len > 53)
					{
						__sto_discA = disc.Substring(0, 53);
						if (len > 139)
						{
							__sto_discB = disc.Substring(53, 86);
							__sto_discC = disc.Substring(139);
						}
						else __sto_discB = disc.Substring(53);
					}
					else __sto_discA = disc;

					break;
				case 2:
					__sto_wtyp = resCommon.__sto_discC;
					__sto_widA = 82.2047F;
					__sto_widB = 410.45669F;

					if (len > 66)
					{
						__sto_discA = disc.Substring(0, 66);
						if (len > 152)
						{
							__sto_discB = disc.Substring(66, 86);
							__sto_discC = disc.Substring(152);
						}
						else __sto_discB = disc.Substring(66);
					}
					else __sto_discA = disc;

					break;
				default:
					__sto_wtyp = resCommon.__sto_discA;
					__sto_widA = 303.02362F;
					__sto_widB = 189.6378F;

					if (len > 64)
					{
						__sto_discA = disc.Substring(0, 64);
						if (len > 150)
						{
							__sto_discB = disc.Substring(64, 86);
							__sto_discC = disc.Substring(150);
						}
						else __sto_discB = disc.Substring(64);
					}
					else __sto_discA = disc;

					break;
			}

			double __sto_spacingA = 
				Convert.ToDouble(resDocument.__sto_spacingA);
			double __sto_spacingB =
				Convert.ToDouble(resDocument.__sto_spacingB);

			Word.WdParagraphAlignment alignCenter = 
				Word.WdParagraphAlignment.wdAlignParagraphCenter;
			Word.WdParagraphAlignment alignLeft = 
				Word.WdParagraphAlignment.wdAlignParagraphLeft;
			Word.WdParagraphAlignment alignJustify = 
				Word.WdParagraphAlignment.wdAlignParagraphJustify;

			Word.WdBorderType bdTop = Word.WdBorderType.wdBorderTop;
			Word.WdBorderType bdBottom = Word.WdBorderType.wdBorderBottom;

			Word.WdLineStyle __sto_line = Word.WdLineStyle.wdLineStyleSingle;

			app = new Word.Application();
			app.Visible = false;

			doc = app.Documents.Add
			(
				ref oMissing, 
				ref oMissing, 
				ref oMissing, 
				ref oMissing
			);

			doc.PageSetup.LeftMargin = 
				Convert.ToSingle(resDocument.__sto_indentL);
			doc.PageSetup.RightMargin =
				Convert.ToSingle(resDocument.__sto_indentR);
			doc.PageSetup.TopMargin =
				Convert.ToSingle(resDocument.__sto_indentT);
			doc.PageSetup.BottomMargin =
				Convert.ToSingle(resDocument.__sto_indentB);

			// Page processing block:
			Word.Paragraph para;
			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = (__sto_size - 2);
			para.Range.Font.AllCaps = 1;
			para.Range.Text = resCommon.__ministry;
			para.Format.Alignment = alignCenter;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = (__sto_size - 1);
			para.Range.Font.AllCaps = 0;
			para.Range.Text = resCommon.__universityA;
			para.Format.Alignment = alignCenter;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = (__sto_size - 1);
			para.Range.Font.AllCaps = 0;
			para.Range.Text = resCommon.__universityB;
			para.Format.Alignment = alignCenter;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = __sto_size;
			para.Range.Font.Bold = 1;
			para.Range.Text = resCommon.__universityC;
			para.Format.Alignment = alignCenter;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = (__sto_size + 1);
			para.Range.Font.Bold = 0;
			para.Range.Text = __sto_null;
			para.Format.Alignment = alignJustify;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = (__sto_size + 1);
			para.Range.Font.Bold = 0;
			para.Range.Text = __sto_null;
			para.Format.Alignment = alignJustify;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = __sto_size;
			para.Range.Font.Bold = 0;
			para.Range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
			para.Range.Text = uHighSchool;
			para.Format.Alignment = alignCenter;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = __sto_size;
			para.Range.Font.Superscript = 1;
			para.Range.Text = resCommon.__sto_hs;
			para.Format.Alignment = alignCenter;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = (__sto_size + 1);
			para.Range.Font.Superscript = 0;
			para.Range.Text = __sto_null;
			para.Format.Alignment = alignCenter;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = (__sto_size + 3);
			para.Range.Font.Bold = 1;
			para.Range.Font.AllCaps = 1;
			para.Range.Text = __sto_work;
			para.Format.Alignment = alignCenter;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = (__sto_size + 1);
			para.Range.Font.Bold = 0;
			para.Range.Text = __sto_null;
			para.Format.Alignment = alignJustify;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = (__sto_size + 1);
			para.Range.Font.Bold = 0;
			para.Range.Text = __sto_null;
			para.Format.Alignment = alignJustify;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Range.InsertParagraphAfter();

			Word.Table table;
			Word.Range range = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
			table = doc.Content.Tables.Add
			(
				range, 
				5, 2, 
				ref oMissing, ref oMissing
			);
			table.PreferredWidthType = 
				Word.WdPreferredWidthType.wdPreferredWidthPoints;
			table.PreferredWidth = 492.66142F;
			table.Range.Font.Name = __sto_font;
			table.Range.Font.Size = __sto_size;
			table.Range.Font.Bold = 0;
			table.Range.Font.AllCaps = 0;
			table.Range.ParagraphFormat.SpaceAfter = 0;

			table.Cell(1, 1).Range.Text = __sto_wtyp;
			table.Cell(1, 1).Height = 13.6063F;
			table.Cell(1, 1).Width = __sto_widA;
			table.Cell(1, 2).Width = __sto_widB;
			table.Cell(1, 2).Range.Text = __sto_discA;
			table.Cell(1, 2).Range.ParagraphFormat.Alignment = alignCenter;

			table.Cell(2, 2).Merge(table.Cell(2, 1));
			table.Cell(2, 1).Height = 13.6063F;
			table.Cell(2, 1).Width = 492.66142F;
			table.Cell(2, 1).Range.Text = __sto_discB;
			table.Cell(2, 1).Range.ParagraphFormat.Alignment = alignJustify;
			table.Cell(2, 1).Range.Borders[bdBottom].LineStyle = __sto_line;

			table.Cell(3, 2).Merge(table.Cell(3, 1));
			table.Cell(3, 1).Height = 13.6063F;
			table.Cell(3, 1).Width = 492.66142F;
			table.Cell(3, 1).Range.Text = __sto_discC;
			table.Cell(3, 1).Range.ParagraphFormat.Alignment = alignJustify;
			table.Cell(3, 1).Range.Borders[bdBottom].LineStyle = __sto_line;

			table.Cell(4, 1).Range.Text = "\n" + resCommon.__sto_theme;
			table.Cell(4, 1).Height = 27.77953F;
			table.Cell(4, 1).Width = 54.99213F;
			table.Cell(4, 2).Width = 437.66929F;
			table.Cell(4, 2).Range.Borders[bdBottom].LineStyle = __sto_line;
			table.Cell(4, 2).Range.Text = dTheme;
			table.Cell(4, 2).Range.ParagraphFormat.Alignment = alignCenter;
			table.Cell(4, 2).VerticalAlignment =
				Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom;

			table.Cell(5, 2).Merge(table.Cell(5, 1));
			table.Cell(5, 1).Range.Font.Size = (__sto_size + 1);
			table.Cell(5, 1).Range.Text = " \n";
			table.Cell(5, 1).Width = 492.66142F;
			table.Cell(1, 2).Range.Borders[bdBottom].LineStyle = __sto_line;

			range = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
			table = doc.Content.Tables.Add
			(
				range,
				8, 2,
				ref oMissing, ref oMissing
			);

			table.Range.Font.Name = __sto_font;

			for (int i = 6; i < 13; i++)
			{
				table.Cell(i, 1).Width = 203.81F;
				table.Cell(i, 2).Width = 283.46F;
				table.Cell(i, 1).Range.ParagraphFormat.LeftIndent = 35.34F;
			}

			table.Cell(13, 1).Width = 203.81F;
			table.Cell(13, 2).Width = 283.46F;

			table.Cell(6, 2).Range.Font.Size = __sto_size;
			table.Cell(6, 2).Range.Text = __sto_exec + '\n' +
				uSurname + ' ' + uForename + 
				' ' + uPatronymic;
			table.Cell(6, 2).Range.ParagraphFormat.LineSpacingRule = 
				Word.WdLineSpacing.wdLineSpaceSingle;

			table.Cell(7, 2).Range.Font.Superscript = 1;
			table.Cell(7, 2).Range.ParagraphFormat.Alignment = alignCenter;
			table.Cell(7, 2).Range.Borders[bdTop].LineStyle = __sto_line;
			table.Cell(7, 2).Range.Text = resCommon.__sto_stud;
			table.Cell(7, 2).Height = 13.6063F;

			table.Cell(8, 2).Range.Font.Size = __sto_size;
			table.Cell(8, 2).Range.Font.Superscript = 0;
			table.Cell(8, 2).Range.Text = 
				resCommon.__sto_dirA + '\n' + 
				uCode + ' ' + uDirection;
			table.Cell(8, 2).Range.ParagraphFormat.LineSpacingRule = 
				Word.WdLineSpacing.wdLineSpaceSingle;
			table.Cell(8, 2).Range.Borders[bdBottom].LineStyle = __sto_line;

			table.Cell(9, 2).Range.Font.Size = __sto_size;
			table.Cell(9, 2).Range.Font.Superscript = 1;
			table.Cell(9, 2).Range.ParagraphFormat.Alignment = alignCenter;
			table.Cell(9, 2).Range.Text = resCommon.__sto_dirB;
			table.Cell(9, 2).Height = 13.6063F;

			table.Cell(10, 2).Range.Font.Size = __sto_size;
			table.Cell(10, 2).Range.Font.Superscript = 0;
			table.Cell(10, 2).Range.ParagraphFormat.LineSpacingRule = 
				Word.WdLineSpacing.wdLineSpaceSingle;
			table.Cell(10, 2).VerticalAlignment = 
				Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom;
			table.Cell(10, 2).Range.Borders[bdBottom].LineStyle = __sto_line;
			table.Cell(10, 2).Range.Text = 
				resCommon.__sto_course + ' ' + uCourse;
			table.Cell(10, 2).Height = 13.6063F;

			table.Cell(11, 2).Range.Font.Size = __sto_size;
			table.Cell(11, 2).Range.ParagraphFormat.LineSpacingRule =
				Word.WdLineSpacing.wdLineSpaceSingle;
			table.Cell(11, 2).Range.Borders[bdBottom].LineStyle = __sto_line;
			table.Cell(11, 2).VerticalAlignment = 
				Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
			table.Cell(11, 2).Range.Text = 
				resCommon.__sto_group + ' ' + uGroup;
			table.Cell(11, 2).Height = 13.6063F;

			table.Cell(12, 2).Range.Font.Size = __sto_size;
			table.Cell(12, 2).Range.ParagraphFormat.LineSpacingRule =
				Word.WdLineSpacing.wdLineSpaceSingle;
			table.Cell(12, 2).Range.Borders[bdBottom].LineStyle = __sto_line;
			table.Cell(12, 2).VerticalAlignment = 
				Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
			table.Cell(12, 2).Range.Text = '\n' + resCommon.__sto_advA + ":\n" +
				dPrepod + ", " + dPrepodInfo;

			table.Cell(13, 1).Range.Font.Size = __sto_size;
			table.Cell(13, 1).Range.ParagraphFormat.Alignment = alignLeft;
			table.Cell(13, 1).Range.Text = " \n \n";

			table.Cell(13, 2).Range.Font.Superscript = 1;
			table.Cell(13, 2).Range.Font.Size = __sto_size;
			table.Cell(13, 2).Range.ParagraphFormat.Alignment = alignCenter;
			table.Cell(13, 2).Range.Text = resCommon.__sto_prepod;

			range = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
			table = doc.Content.Tables.Add
			(
				range, 
				4, 5, 
				ref oMissing, ref oMissing
			);
			table.Rows.LeftIndent = 0.0F;
			table.Range.Font.Name = __sto_font;

			table.Cell(15, 3).Range.Borders[bdTop].LineStyle = __sto_line;
			table.Cell(15, 5).Range.Borders[bdTop].LineStyle = __sto_line;

			for (int i = 14; i < 18; i++)
			{
				table.Cell(i, 1).Height = 13.6063F;
				table.Cell(i, 1).Width = 118.77F;
				table.Cell(i, 2).Width = 14.17F;
				table.Cell(i, 3).Width = 177.17F;
				table.Cell(i, 4).Width = 21.26F;
				table.Cell(i, 5).Width = 155.91F;
			}

			table.Cell(14, 1).Range.Font.Size = __sto_size;
			table.Cell(14, 1).Range.Text = resCommon.__sto_markA;

			table.Cell(16, 1).Range.Font.Size = __sto_size;
			table.Cell(16, 1).Range.Text = resCommon.__sto_advA;

			table.Cell(15, 4).Range.Font.Size = __sto_size;
			table.Cell(15, 4).Range.Font.Superscript = 1;
			table.Cell(15, 4).Range.ParagraphFormat.Alignment = alignCenter;
			table.Cell(15, 4).Range.Text = __sto_null;

			table.Cell(16, 4).Range.Font.Size = __sto_size;
			table.Cell(16, 4).Range.Font.Superscript = 1;
			table.Cell(16, 4).Range.ParagraphFormat.Alignment = alignCenter;
			table.Cell(16, 4).Range.Text = __sto_null;

			table.Cell(17, 4).Range.Font.Size = __sto_size;
			table.Cell(17, 4).Range.Font.Superscript = 1;
			table.Cell(17, 4).Range.ParagraphFormat.Alignment = alignCenter;
			table.Cell(17, 4).Range.Text = __sto_null;

			table.Cell(15, 3).Range.Font.Size = __sto_size;
			table.Cell(15, 3).Range.Font.Superscript = 1;
			table.Cell(15, 3).Range.ParagraphFormat.Alignment = alignCenter;
			table.Cell(15, 3).Range.Text = resCommon.__sto_markB;

			table.Cell(15, 5).Range.Font.Size = __sto_size;
			table.Cell(15, 5).Range.Font.Superscript = 1;
			table.Cell(15, 5).Range.ParagraphFormat.Alignment = alignCenter;
			table.Cell(15, 5).Range.Text = resCommon.__sto_date;

			table.Cell(17, 3).Range.Font.Size = __sto_size;
			table.Cell(17, 3).Range.Font.Superscript = 1;
			table.Cell(17, 3).Range.ParagraphFormat.Alignment = alignCenter;
			table.Cell(17, 3).Range.Text = resCommon.__sto_advB;

			table.Cell(16, 5).Range.Font.Size = __sto_size;
			table.Cell(16, 5).Range.Font.Superscript = 0;
			table.Cell(16, 5).Range.ParagraphFormat.Alignment = alignCenter;
			table.Cell(16, 5).Range.Text = dPrepodIniz;

			table.Cell(17, 5).Range.Font.Size = __sto_size;
			table.Cell(17, 5).Range.Font.Superscript = 1;
			table.Cell(17, 5).Range.ParagraphFormat.Alignment = alignCenter;
			table.Cell(17, 5).Range.Text = resCommon.__sto_advC;

			table.Cell(16, 5).Range.Borders[bdBottom].LineStyle = __sto_line;
			table.Cell(16, 3).Range.Borders[bdBottom].LineStyle = __sto_line;

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = __sto_size;
			para.Range.Font.Superscript = 0;
			para.Range.Text = __sto_null;
			para.Format.Alignment = alignCenter;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Format.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = __sto_size;
			para.Range.Font.Superscript = 0;
			para.Range.Text = __sto_null;
			para.Format.Alignment = alignCenter;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Format.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = __sto_size;
			para.Range.Font.Superscript = 0;
			para.Range.Text = resUser.__city + ' ' + dYear;
			para.Format.Alignment = alignCenter;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Range.InsertParagraphAfter();

			object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
			object oPageBreak = Word.WdBreakType.wdSectionBreakNextPage;

			range = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
			range.Collapse(ref oCollapseEnd);
			range.InsertBreak(ref oPageBreak);

			Console.WriteLine
			(
				"::title.mktg_o > task completed."
			);

			// REMOVE AFTER OPTIMIZATION!!!!!!!!!!!!!!!!!!!
			PreloadStyles();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			object __sto_sh1 = StyleSet.__sto_sh1;
			para.set_Style(__sto_sh1);
			para.Range.Text = "Лист для замечаний";
			para.Range.InsertParagraphAfter();
		}

		// Makes title page from group 1n.
		// __sto_work :: type of work (course work / course project)
		// group	  :: type of work (group / single)
		// Returns nothing. Launching Word application.
		static void MakeTitleGroupTwo(string __sto_work, int __sto_type = 0)
		{
			Console.WriteLine
			(
				"::title.mktg_t > working"
			);

			GetRegistryKeys();

			// Preparing block:
			string __sto_font = resDocument.__sto_font;
			string __sto_null = " ";
			int __sto_size = Convert.ToInt32(resDocument.__sto_size);

			string __sto_discA = "";
			string __sto_discB = "";
			string __sto_discC = "";

			string __sto_themA = "";
			string __sto_themB = "";

			string __sto_wtyp;
			float __sto_widA;
			float __sto_widB;

			string __sto_exec;

			var gender = Convert.ToBoolean(uGender);
			if (gender) __sto_exec = resCommon.__sto_execM;
			else __sto_exec = resCommon.__sto_execF;

			var disc = dDiscipline;
			var len = disc.Length;
			switch (__sto_type)
			{
				case 0:
					__sto_wtyp = resCommon.__sto_discA;
					__sto_widA = 99.2126F;
					__sto_widB = 388.06299F;

					if (len > 64)
					{
						__sto_discA = disc.Substring(0, 64);
						if (len > 150)
						{
							__sto_discB = disc.Substring(64, 86);
							__sto_discC = disc.Substring(150);
						}
						else __sto_discB = disc.Substring(64);
					}
					else __sto_discA = disc;

					break;
				case 1:
					__sto_wtyp = resCommon.__sto_discB;
					__sto_widA = 195.591F;
					__sto_widB = 297.07087F;

					if (len > 53)
					{
						__sto_discA = disc.Substring(0, 53);
						if (len > 139)
						{
							__sto_discB = disc.Substring(53, 86);
							__sto_discC = disc.Substring(139);
						}
						else __sto_discB = disc.Substring(53);
					}
					else __sto_discA = disc;

					break;
				case 2:
					__sto_wtyp = resCommon.__sto_discC;
					__sto_widA = 82.2047F;
					__sto_widB = 410.45669F;

					if (len > 66)
					{
						__sto_discA = disc.Substring(0, 66);
						if (len > 152)
						{
							__sto_discB = disc.Substring(66, 86);
							__sto_discC = disc.Substring(152);
						}
						else __sto_discB = disc.Substring(66);
					}
					else __sto_discA = disc;

					break;
				default:
					__sto_wtyp = resCommon.__sto_discA;
					__sto_widA = 303.02362F;
					__sto_widB = 189.6378F;

					if (len > 64)
					{
						__sto_discA = disc.Substring(0, 64);
						if (len > 150)
						{
							__sto_discB = disc.Substring(64, 86);
							__sto_discC = disc.Substring(150);
						}
						else __sto_discB = disc.Substring(64);
					}
					else __sto_discA = disc;

					break;
			}

			double __sto_spacingA =
				Convert.ToDouble(resDocument.__sto_spacingA);
			double __sto_spacingB =
				Convert.ToDouble(resDocument.__sto_spacingB);

			Word.WdParagraphAlignment alignCenter =
				Word.WdParagraphAlignment.wdAlignParagraphCenter;
			Word.WdParagraphAlignment alignLeft =
				Word.WdParagraphAlignment.wdAlignParagraphLeft;
			Word.WdParagraphAlignment alignJustify =
				Word.WdParagraphAlignment.wdAlignParagraphJustify;

			Word.WdBorderType bdTop = Word.WdBorderType.wdBorderTop;
			Word.WdBorderType bdBottom = Word.WdBorderType.wdBorderBottom;

			Word.WdLineStyle __sto_line = Word.WdLineStyle.wdLineStyleSingle;

			app = new Word.Application();
			app.Visible = false;

			doc = app.Documents.Add
			(
				ref oMissing,
				ref oMissing,
				ref oMissing,
				ref oMissing
			);

			doc.PageSetup.LeftMargin =
				Convert.ToSingle(resDocument.__sto_indentL);
			doc.PageSetup.RightMargin =
				Convert.ToSingle(resDocument.__sto_indentR);
			doc.PageSetup.TopMargin =
				Convert.ToSingle(resDocument.__sto_indentT);
			doc.PageSetup.BottomMargin =
				Convert.ToSingle(resDocument.__sto_indentB);

			if (dTheme.Length > 70)
			{
				__sto_themA = dTheme.Substring(0, 70);
				__sto_themB = dTheme.Substring(70);
			}
			else __sto_themA = dTheme;

			// Page processing block:
			Word.Paragraph para;
			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = (__sto_size - 2);
			para.Range.Font.AllCaps = 1;
			para.Range.Text = resCommon.__ministry;
			para.Format.Alignment = alignCenter;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = (__sto_size - 1);
			para.Range.Font.AllCaps = 0;
			para.Range.Text = resCommon.__universityA;
			para.Format.Alignment = alignCenter;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = (__sto_size - 1);
			para.Range.Font.AllCaps = 0;
			para.Range.Text = resCommon.__universityB;
			para.Format.Alignment = alignCenter;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = __sto_size;
			para.Range.Font.Bold = 1;
			para.Range.Text = resCommon.__universityC;
			para.Format.Alignment = alignCenter;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = (__sto_size - 1);
			para.Range.Font.Bold = 0;
			para.Range.Text = __sto_null;
			para.Format.Alignment = alignJustify;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = __sto_size;
			para.Range.Font.Bold = 0;
			para.Range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
			para.Range.Text = uHighSchool;
			para.Format.Alignment = alignCenter;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = __sto_size;
			para.Range.Font.Superscript = 1;
			para.Range.Text = resCommon.__sto_hs;
			para.Format.Alignment = alignCenter;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = (__sto_size + 3);
			para.Range.Font.Superscript = 0;
			para.Range.Text = __sto_null;
			para.Format.Alignment = alignCenter;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = (__sto_size + 3);
			para.Range.Font.Bold = 1;
			para.Range.Font.AllCaps = 1;
			para.Range.Text = __sto_work;
			para.Format.Alignment = alignCenter;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = (__sto_size + 3);
			para.Range.Font.Bold = 0;
			para.Range.Text = __sto_null;
			para.Format.Alignment = alignJustify;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Range.InsertParagraphAfter();

			Word.Table table;
			Word.Range range = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
			table = doc.Content.Tables.Add
			(
				range,
				5, 2,
				ref oMissing, ref oMissing
			);
			table.PreferredWidthType =
				Word.WdPreferredWidthType.wdPreferredWidthPoints;
			table.PreferredWidth = 487.27559F;
			table.Range.Font.Name = __sto_font;
			table.Range.Font.Size = __sto_size;
			table.Range.Font.Bold = 0;
			table.Range.Font.AllCaps = 0;
			table.Range.ParagraphFormat.SpaceAfter = 0;

			table.Cell(1, 1).Range.Text = __sto_wtyp;
			table.Cell(1, 1).Height = 13.6063F;
			table.Cell(1, 1).Width = __sto_widA;
			table.Cell(1, 2).Width = __sto_widB;
			table.Cell(1, 2).Range.Text = __sto_discA;
			table.Cell(1, 2).Range.ParagraphFormat.Alignment = alignCenter;

			table.Cell(2, 2).Merge(table.Cell(2, 1));
			table.Cell(2, 1).Height = 13.6063F;
			table.Cell(2, 1).Width = 487.27559F;
			table.Cell(2, 1).Range.Text = __sto_discB;
			table.Cell(2, 1).Range.ParagraphFormat.Alignment = alignJustify;
			table.Cell(2, 1).Range.Borders[bdBottom].LineStyle = __sto_line;

			table.Cell(3, 2).Merge(table.Cell(3, 1));
			table.Cell(3, 1).Height = 13.6063F;
			table.Cell(3, 1).Width = 487.27559F;
			table.Cell(3, 1).Range.Text = __sto_discC;
			table.Cell(3, 1).Range.ParagraphFormat.Alignment = alignJustify;
			table.Cell(3, 1).Range.Borders[bdBottom].LineStyle = __sto_line;

			table.Cell(4, 1).Range.Text = "\n" + resCommon.__sto_theme;
			table.Cell(4, 1).Height = 27.77953F;
			table.Cell(4, 1).Width = 62.07874F;
			table.Cell(4, 2).Width = 425.197F;
			table.Cell(4, 2).Range.Borders[bdBottom].LineStyle = __sto_line;
			table.Cell(4, 2).Range.Text = __sto_themA;
			table.Cell(4, 2).Range.ParagraphFormat.Alignment = alignCenter;
			table.Cell(4, 2).VerticalAlignment =
				Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom;

			table.Cell(5, 2).Merge(table.Cell(5, 1));
			table.Cell(5, 1).Range.Font.Size = __sto_size;
			table.Cell(5, 1).Range.Text = __sto_themB;
			table.Cell(5, 1).Range.ParagraphFormat.Alignment = alignJustify;
			table.Cell(5, 1).Width = 487.27559F;

			table.Cell(5, 1).Range.Borders[bdBottom].LineStyle = __sto_line;
			table.Cell(1, 2).Range.Borders[bdBottom].LineStyle = __sto_line;

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = (__sto_size + 3);
			para.Range.Font.Bold = 1;
			para.Range.Text = __sto_null;
			para.Format.Alignment = alignCenter;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = (__sto_size + 3);
			para.Range.Font.Bold = 1;
			para.Range.Text = __sto_null;
			para.Format.Alignment = alignJustify;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Range.InsertParagraphAfter();

			range = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
			table = doc.Content.Tables.Add
			(
				range,
				8, 2,
				ref oMissing, ref oMissing
			);

			table.Range.Font.Name = __sto_font;

			for (int i = 1; i < 8; i++)
			{
				table.Cell(i, 1).Width = 203.81F;
				table.Cell(i, 2).Width = 283.46F;
				table.Cell(i, 1).Range.ParagraphFormat.LeftIndent = 35.34F;
			}

			table.Cell(8, 1).Width = 203.81F;
			table.Cell(8, 2).Width = 283.46F;

			table.Cell(1, 1).Range.Font.Size = __sto_size;
			table.Cell(1, 1).Range.Text = " \n";

			table.Cell(1, 2).Range.Font.Size = __sto_size;
			table.Cell(1, 2).Range.Text = __sto_exec + '\n' +
				uSurname + ' ' + uForename +
				' ' + uPatronymic;
			table.Cell(1, 2).Range.ParagraphFormat.LineSpacingRule =
				Word.WdLineSpacing.wdLineSpaceSingle;

			table.Cell(2, 2).Range.Font.Superscript = 1;
			table.Cell(2, 2).Range.ParagraphFormat.Alignment = alignCenter;
			table.Cell(2, 2).Range.Borders[bdTop].LineStyle = __sto_line;
			table.Cell(2, 2).Range.Text = resCommon.__sto_stud;
			table.Cell(2, 2).Height = 13.6063F;

			table.Cell(3, 2).Range.Font.Size = __sto_size;
			table.Cell(3, 2).Range.Font.Superscript = 0;
			table.Cell(3, 2).Range.Text =
				resCommon.__sto_dirA + '\n' +
				uCode + ' ' + uDirection;
			table.Cell(3, 2).Range.ParagraphFormat.LineSpacingRule =
				Word.WdLineSpacing.wdLineSpaceSingle;
			table.Cell(3, 2).Range.Borders[bdBottom].LineStyle = __sto_line;

			table.Cell(4, 2).Range.Font.Size = __sto_size;
			table.Cell(4, 2).Range.Font.Superscript = 1;
			table.Cell(4, 2).Range.ParagraphFormat.Alignment = alignCenter;
			table.Cell(4, 2).Range.Text = resCommon.__sto_dirB;
			table.Cell(4, 2).Height = 13.6063F;

			table.Cell(5, 2).Range.Font.Size = __sto_size;
			table.Cell(5, 2).Range.Font.Superscript = 0;
			table.Cell(5, 2).Range.ParagraphFormat.LineSpacingRule =
				Word.WdLineSpacing.wdLineSpaceSingle;
			table.Cell(5, 2).VerticalAlignment =
				Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom;
			table.Cell(5, 2).Range.Borders[bdBottom].LineStyle = __sto_line;
			table.Cell(5, 2).Range.Text =
				resCommon.__sto_course + ' ' + uCourse;
			table.Cell(5, 2).Height = 13.6063F;

			table.Cell(6, 2).Range.Font.Size = __sto_size;
			table.Cell(6, 2).Range.ParagraphFormat.LineSpacingRule =
				Word.WdLineSpacing.wdLineSpaceSingle;
			table.Cell(6, 2).Range.Borders[bdBottom].LineStyle = __sto_line;
			table.Cell(6, 2).VerticalAlignment =
				Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
			table.Cell(6, 2).Range.Text =
				resCommon.__sto_group + ' ' + uGroup;
			table.Cell(6, 2).Height = 13.6063F;

			table.Cell(7, 2).Range.Font.Size = __sto_size;
			table.Cell(7, 2).Range.ParagraphFormat.LineSpacingRule =
				Word.WdLineSpacing.wdLineSpaceSingle;
			table.Cell(7, 2).Range.Borders[bdBottom].LineStyle = __sto_line;
			table.Cell(7, 2).VerticalAlignment =
				Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
			table.Cell(7, 2).Range.Text = '\n' + resCommon.__sto_advA + ":\n" +
				dPrepod + ", " + dPrepodInfo;

			table.Cell(8, 1).Range.Font.Size = __sto_size;
			table.Cell(8, 1).Range.ParagraphFormat.Alignment = alignLeft;
			table.Cell(8, 1).Range.Text = " \n \n";

			table.Cell(8, 2).Range.Font.Superscript = 1;
			table.Cell(8, 2).Range.Font.Size = __sto_size;
			table.Cell(8, 2).Range.ParagraphFormat.Alignment = alignCenter;
			table.Cell(8, 2).Range.Text = resCommon.__sto_prepod;

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = (__sto_size + 3);
			para.Range.Font.Bold = 0;
			para.Range.Text = __sto_null;
			para.Format.Alignment = alignJustify;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = (__sto_size + 2);
			para.Range.Font.Bold = 0;
			para.Range.Text = __sto_null;
			para.Format.Alignment = alignJustify;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Format.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = (__sto_size + 2);
			para.Range.Font.Bold = 0;
			para.Range.Text = __sto_null;
			para.Format.Alignment = alignJustify;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Format.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
			para.Range.InsertParagraphAfter();

			range = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
			table = doc.Content.Tables.Add
			(
				range,
				4, 5,
				ref oMissing, ref oMissing
			);
			table.Rows.LeftIndent = 0.0F;
			table.Range.Font.Name = __sto_font;
			table.Range.ParagraphFormat.LineSpacingRule = 
				Word.WdLineSpacing.wdLineSpaceSingle;

			table.Cell(2, 3).Range.Borders[bdTop].LineStyle = __sto_line;
			table.Cell(2, 5).Range.Borders[bdTop].LineStyle = __sto_line;

			for (int i = 1; i < 5; i++)
			{
				table.Cell(i, 1).Height = 27.77953F;
				table.Cell(i, 2).Height = 13.6063F;
				table.Cell(i, 3).Height = 13.6063F;
				table.Cell(i, 4).Height = 13.6063F;
				table.Cell(i, 5).Height = 13.6063F;

				table.Cell(i, 1).Width = 189.6378F;
				table.Cell(i, 2).Width = 14.1732F;
				table.Cell(i, 3).Width = 141.732F;
				table.Cell(i, 4).Width = 14.1732F;
				table.Cell(i, 5).Width = 127.559F;
			}

			table.Cell(1, 1).Range.Font.Size = __sto_size;
			table.Cell(1, 1).Range.Text = resCommon.__sto_markC;

			table.Cell(3, 1).Range.Font.Size = __sto_size;
			table.Cell(3, 1).Range.Text = resCommon.__sto_advA;

			table.Cell(2, 4).Range.Font.Size = __sto_size;
			table.Cell(2, 4).Range.Font.Superscript = 1;
			table.Cell(2, 4).Range.ParagraphFormat.Alignment = alignCenter;
			table.Cell(2, 4).Range.Text = __sto_null;

			table.Cell(3, 4).Range.Font.Size = __sto_size;
			table.Cell(3, 4).Range.Font.Superscript = 1;
			table.Cell(3, 4).Range.ParagraphFormat.Alignment = alignCenter;
			table.Cell(3, 4).Range.Text = __sto_null;

			table.Cell(4, 4).Range.Font.Size = __sto_size;
			table.Cell(4, 4).Range.Font.Superscript = 1;
			table.Cell(4, 4).Range.ParagraphFormat.Alignment = alignCenter;
			table.Cell(4, 4).Range.Text = __sto_null;

			table.Cell(2, 3).Range.Font.Size = __sto_size;
			table.Cell(2, 3).Range.Font.Superscript = 1;
			table.Cell(2, 3).Range.ParagraphFormat.Alignment = alignCenter;
			table.Cell(2, 3).Range.Text = resCommon.__sto_markB;

			table.Cell(2, 5).Range.Font.Size = __sto_size;
			table.Cell(2, 5).Range.Font.Superscript = 1;
			table.Cell(2, 5).Range.ParagraphFormat.Alignment = alignCenter;
			table.Cell(2, 5).Range.Text = resCommon.__sto_date;

			table.Cell(4, 3).Range.Font.Size = __sto_size;
			table.Cell(4, 3).Range.Font.Superscript = 1;
			table.Cell(4, 3).Range.ParagraphFormat.Alignment = alignCenter;
			table.Cell(4, 3).Range.Text = resCommon.__sto_advB;

			table.Cell(3, 5).Range.Font.Size = __sto_size;
			table.Cell(3, 5).Range.Font.Superscript = 0;
			table.Cell(3, 5).Range.ParagraphFormat.Alignment = alignCenter;
			table.Cell(3, 5).Range.Text = dPrepodIniz;

			table.Cell(4, 5).Range.Font.Size = __sto_size;
			table.Cell(4, 5).Range.Font.Superscript = 1;
			table.Cell(4, 5).Range.ParagraphFormat.Alignment = alignCenter;
			table.Cell(4, 5).Range.Text = resCommon.__sto_advC;

			table.Cell(3, 5).Range.Borders[bdBottom].LineStyle = __sto_line;
			table.Cell(3, 3).Range.Borders[bdBottom].LineStyle = __sto_line;

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = (__sto_size - 1);
			para.Range.Font.Superscript = 0;
			para.Range.Text = __sto_null;
			para.Format.Alignment = alignJustify;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = (__sto_size - 1);
			para.Range.Font.Superscript = 0;
			para.Range.Text = __sto_null;
			para.Format.Alignment = alignJustify;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = (__sto_size + 1);
			para.Range.Font.Superscript = 0;
			para.Range.Text = __sto_null;
			para.Format.Alignment = alignCenter;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Format.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = (__sto_size + 1);
			para.Range.Font.Superscript = 0;
			para.Range.Text = __sto_null;
			para.Format.Alignment = alignCenter;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Format.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
			para.Range.InsertParagraphAfter();

			para = doc.Content.Paragraphs.Add(ref oMissing);
			para.Range.Font.Name = __sto_font;
			para.Range.Font.Size = __sto_size;
			para.Range.Font.Superscript = 0;
			para.Range.Text = resUser.__city + ' ' + dYear;
			para.Format.Alignment = alignCenter;
			para.Format.SpaceAfter = 0;
			para.Format.SpaceBefore = 0;
			para.Range.InsertParagraphAfter();

			object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
			object oPageBreak = Word.WdBreakType.wdPageBreak;

			range = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
			range.Collapse(ref oCollapseEnd);
			range.InsertBreak(ref oPageBreak);

			Console.WriteLine
			(
				"::title.mktg_t > task completed."
			);
		}

		static void AttachNotelist()
        {
			
        }

		static void InsertContent()
        {

        }

		static void PreloadStyles()
        {
			Word.Style sh1 = doc.Styles.Add
			(
				StyleSet.__sto_sh1,
				Word.WdStyleType.wdStyleTypeParagraph
			);
			sh1.Font.Name = "Times New Roman";
			sh1.Font.Size = 12.0F;
			sh1.Font.Bold = 1;
			sh1.Font.AllCaps = 1;
			sh1.ParagraphFormat.LeftIndent = ConstantIndents.__1p25;
			sh1.ParagraphFormat.RightIndent = ConstantIndents.__2p50;
			sh1.ParagraphFormat.SpaceAfter = 0.0F;
			sh1.ParagraphFormat.SpaceBefore = 6.0F;
			sh1.ParagraphFormat.Alignment = 
				Word.WdParagraphAlignment.wdAlignParagraphJustify;

			Word.Style sh2 = doc.Styles.Add
			(
				StyleSet.__sto_sh2,
				Word.WdStyleType.wdStyleTypeParagraph
			);
			sh2.Font.Name = "Times New Roman";
			sh2.Font.Size = 12.0F;
			sh2.Font.Bold = 0;
			sh2.ParagraphFormat.LeftIndent = ConstantIndents.__1p25;
			sh2.ParagraphFormat.SpaceAfter = 12.0F;
			sh2.ParagraphFormat.SpaceBefore = 12.0F;
			sh2.ParagraphFormat.Alignment =
				Word.WdParagraphAlignment.wdAlignParagraphCenter;

			Word.Style dh1 = doc.Styles.Add
			(
				StyleSet.__sto_dh1,
				Word.WdStyleType.wdStyleTypeParagraph
			);
			dh1.Font.Name = "Times New Roman";
			dh1.Font.Size = 12.0F;
			dh1.Font.Bold = 0;
			dh1.Font.AllCaps = 1;
			dh1.ParagraphFormat.LeftIndent = ConstantIndents.__1p00;
			dh1.ParagraphFormat.SpaceAfter = 12.0F;
			dh1.ParagraphFormat.SpaceBefore = 0.0F;
			dh1.ParagraphFormat.Alignment =
				Word.WdParagraphAlignment.wdAlignParagraphCenter;

			Word.Style dft = doc.Styles.Add
			(
				StyleSet.__sto_dft,
				Word.WdStyleType.wdStyleTypeParagraph
			);
			dft.Font.Name = "Times New Roman";
			dft.Font.Size = 12.0F;
			dft.Font.Bold = 0;
			dft.ParagraphFormat.LeftIndent = ConstantIndents.__1p25;
			dft.ParagraphFormat.SpaceAfter = 0.0F;
			dft.ParagraphFormat.SpaceBefore = 0.0F;
			dft.ParagraphFormat.LineSpacingRule = 
				Word.WdLineSpacing.wdLineSpace1pt5;
			dft.ParagraphFormat.Alignment =
				Word.WdParagraphAlignment.wdAlignParagraphJustify;
		}
	}
}
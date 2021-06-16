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
		static object oMissing = Missing.Value;
		static object oEndOfDoc = "\\endofdoc";
		static object g_style = "GlobalStyle";

		static Word.Application app;
		static Word.Document doc;

		static void Main(string[] args)
		{
			try
			{
				if (args.Length == 0)
				{
					Console.WriteLine("> Invalid args.");
					Help();
				}
				else if (args.Length > 0)
				{
					var cmd = args[0];
					switch (cmd)
					{
						default:
							Console.WriteLine("Unknown args.");
							Help();
							break;
						case "--title" when args.Length == 2:
							int pattern = Convert.ToInt32(args[1]);
							Title(pattern);
							break;
					}

					app.Visible = true;
				}
			}
			catch(Exception e)
			{
				WriteException(e);
				Console.ReadKey();
			}
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
			Console.WriteLine("Runtime error: exception catched.");
			Console.WriteLine("Info:");
			Console.WriteLine(e.Message);
			Console.WriteLine("Task aborted.\n");
		}

		// Generates title from internal pattern.
		static void Title(int pattern)
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

			switch(pattern)
			{
				case 1:
					__sto_work = "Эссе";
					MakeTitleGroupOne(__sto_work);
					break;
				case 2:
					__sto_work = "Реферат";
					MakeTitleGroupOne(__sto_work);
					break;
				case 3:
					__sto_work = "Контрольная работа";
					MakeTitleGroupOne(__sto_work);
					break;
				case 4:
					__sto_work = "Расчётно-графическая работа";
					MakeTitleGroupOne(__sto_work);
					break;
				case 11:
					__sto_work = "Курсовой проект";
					MakeTitleGroupOne(__sto_work);
					break;
				default:
					throw new InvalidOperationException();
			}
		}

		// Makes title page from group 0n
		// Returns nothing. Launching Word application.
		static void MakeTitleGroupOne(string __sto_work, int __sto_type = 0)
		{
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

			var gender = Convert.ToBoolean(resUser.__sto_g);
			if (gender) __sto_exec = resCommon.__sto_execM;
			else __sto_exec = resCommon.__sto_execF;

			var disc = resDocument.__disc;
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

			Word.WdBorderType bdLeft = Word.WdBorderType.wdBorderLeft;
			Word.WdBorderType bdRight = Word.WdBorderType.wdBorderRight;
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
			para.Range.Text = resCommon.__hs;
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
			table.Cell(4, 2).Range.Text = resDocument.__theme;
			table.Cell(4, 2).Range.ParagraphFormat.Alignment = alignCenter;
			table.Cell(4, 2).VerticalAlignment =
				Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom;

			table.Cell(5, 2).Merge(table.Cell(5, 1));
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
				table.Cell(i, 1).Range.ParagraphFormat.LeftIndent = 35.34F;
            }

			table.Cell(13, 1).Width = 203.81F;

			table.Cell(6, 2).Width = 283.46F;
			table.Cell(6, 2).Range.Font.Size = __sto_size;
			table.Cell(6, 2).Range.Text = __sto_exec + '\n' +
				resUser.__surname + ' ' + resUser.__forename + 
				' ' + resUser.__patronymic;
			table.Cell(6, 2).Range.ParagraphFormat.LineSpacingRule = 
				Word.WdLineSpacing.wdLineSpaceSingle;

		}
	}
}

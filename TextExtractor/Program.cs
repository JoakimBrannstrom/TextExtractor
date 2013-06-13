using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace TextExtractor
{
	class Program
	{
		static void Main()
		{
			do
			{
				Console.Clear();

				RunExtraction();

				Console.WriteLine("{0}{0}Press 'r' to run once more.", Environment.NewLine);
			}while(Console.ReadKey().KeyChar == 'r');
		}

		private static void RunExtraction()
		{
			var filesPath = string.Format("{0}\\Files", Environment.CurrentDirectory);
			var dirInfo = new DirectoryInfo(filesPath);

			if (!dirInfo.Exists)
			{
				Console.WriteLine("The path '{0}' did not exist. Press <enter> to exit.", filesPath);
				Console.ReadLine();
				//return;
			}

			var files = dirInfo.EnumerateFiles().ToArray();

			var textExtractor = new TextExtractor();

			ShowFileParsingInfo(textExtractor, files);

            Console.WriteLine("{0}{0}Press 'enter' to continue.", Environment.NewLine);
            Console.ReadLine();

			ParseFiles(textExtractor, files);
		}

		private static void ShowFileParsingInfo(ITextExtractor textExtractor, FileInfo[] files)
		{
			PrintFilesList("Found the following files:", files);

			var timer = new Stopwatch();
			timer.Start();
			var parsableFiles = files.Where(f => textExtractor.IsParseable(f.FullName)).ToArray();
			timer.Stop();
			Console.WriteLine("{0}{0}Processed in {1}{0}{0}", Environment.NewLine, timer.Elapsed);

			PrintFilesList("Parsable files:", parsableFiles);

			var unparsableFiles = files.Where(f => textExtractor.IsParseable(f.FullName) == false).ToArray();
			PrintFilesList("Unparsable files:", unparsableFiles);
		}

		private static void PrintFilesList(string header, IEnumerable<FileInfo> files)
		{
			Console.WriteLine(header);
			Console.WriteLine(string.Join(Environment.NewLine, files.Select(f => "\t- " + f.Name)));
			Console.WriteLine();
		}

		private static void ParseFiles(ITextExtractor textExtractor, IEnumerable<FileInfo> files)
		{
			var parsableFiles = files.Where(f => textExtractor.IsParseable(f.FullName)).ToArray();

			var timer = new Stopwatch();
			timer.Start();

			foreach (var fileInfo in parsableFiles)
			{
				Console.WriteLine("{0}{0}---------------------------------------", Environment.NewLine);
				Console.WriteLine("Parsing file: {0}{1}", fileInfo.Name, Environment.NewLine);

				var result = ParseFile(textExtractor, fileInfo.FullName);
				Console.WriteLine("Parsed result is: ");
				Console.Write(result);
			}

			timer.Stop();
			Console.WriteLine("{0}{0}Processed in {1}", Environment.NewLine, timer.Elapsed);
		}

		private static string ParseFile(ITextExtractor textExtractor, string filename)
		{
			// Images: https://www.google.se/search?q=tiff+ifilter&ie=utf-8&oe=utf-8&aq=t&rls=org.mozilla:en-US:official&client=firefox-a&channel=fflb#hl=sv&client=firefox-a&hs=Cis&tbo=d&rls=org.mozilla:en-US%3Aofficial&channel=fflb&sclient=psy-ab&q=image+ifilter&oq=image+ifilter&gs_l=serp.3..0i13l4.4693.4693.1.4936.1.1.0.0.0.0.55.55.1.1.0...0.0...1c.1.k9PU0a3S1m4&pbx=1&bav=on.2,or.r_gc.r_pw.r_qf.&bvm=bv.1355534169,d.bGE&fp=71792659602ba5ba&bpcl=40096503&biw=1680&bih=919
			// http://technet.microsoft.com/sv-se/library/dd834685.aspx
			// http://technet.microsoft.com/en-us/library/dd744701%28v=ws.10%29.aspx

			// var filename = Environment.CurrentDirectory + "\\test.pdf";   
			// x64, working: http://www.adobe.com/support/downloads/thankyou.jsp?ftpID=4025&fileID=3941
			// http://www.adobe.com/support/downloads/detail.jsp?ftpID=2611
			// http://www.foxitsoftware.com/products/ifilter/

			// http://www.microsoft.com/en-us/download/details.aspx?id=3988
			// var filename = Environment.CurrentDirectory + "\\test.docx";    // http://www.microsoft.com/en-us/download/details.aspx?id=20109
			// http://www.microsoft.com/en-us/download/details.aspx?id=17062
			// var filename = Environment.CurrentDirectory + "\\test.txt";
			/*
			try
			{
				using (var stream = File.OpenRead(filename))
				{
					return textExtractor.GetText(stream);
				}
			}
			catch(Exception exc)
			{
				Console.WriteLine("Exception was thrown while reading '" + filename + "'");
				Console.WriteLine("Exception: " + exc.Message);
				return "";
			}
			*/
			return textExtractor.GetText(filename);
		}
	}
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace TextExtractor
{
	public interface ITextExtractor
	{
		bool IsParseable(string filename);
		string GetText(string path);
		string GetText(Stream stream);
	}

	/// Convenience class which provides static methods to extract text from files using installed IFilters
	public class TextExtractor : ITextExtractor
	{
		public bool IsParseable(string filename)
		{
			var filter = LoadIFilter(filename);

			if (filter == null)
				return false;

			Marshal.ReleaseComObject(filter);
			GC.Collect();
			GC.WaitForPendingFinalizers();

			return true;
		}

		private static IFilter LoadIFilter(string filename)
		{
			IFilter filter = null;

			// Try to load the corresponding IFilter
			if (LoadIFilter(filename, null, ref filter) == (int)FilterReturnCodes.Success)
				return filter;

			return null;
		}

		[DllImport("query.dll", CharSet = CharSet.Unicode)]
		private static extern int LoadIFilter(string pwcsPath, [MarshalAs(UnmanagedType.IUnknown)] object pUnkOuter, ref IFilter ppIUnk);

		[DllImport("query.dll", CharSet = CharSet.Unicode)]
		private static extern int BindIFilterFromStream(IStream pStm, 
														[MarshalAs(UnmanagedType.IUnknown)] object pUnkOuter,
														[Out] out IFilter ppIUnk);

		[DllImport("ole32.dll", CharSet = CharSet.Unicode)]
		private static extern void CreateStreamOnHGlobal(IntPtr hGlobal, bool fDeleteOnRelease, [Out] out IStream pStream);

		private static IFilter LoadIFilter(Stream stream)
		{
			// copy stream to byte array
			var b = new byte[stream.Length];
			stream.Read(b, 0, b.Length);

			// allocate space on the native heap
			var nativePtr = Marshal.AllocHGlobal(b.Length);
	
			// copy byte array to native heap
			Marshal.Copy(b, 0, nativePtr, b.Length);

			// Create a UCOMIStream from the allocated memory
			IStream comStream;

			CreateStreamOnHGlobal(nativePtr, true, out comStream);

			// Try to load the corresponding IFilter 
			IFilter filter;
			int resultLoad = BindIFilterFromStream(comStream, null, out filter);
			if (resultLoad == (int)FilterReturnCodes.Success)
				return filter;

			Marshal.ThrowExceptionForHR(resultLoad);
			return null;
		}

		public string GetText(Stream stream)
		{
			IFilter filter = null;

			try
			{
				filter = LoadIFilter(stream);

				if (filter == null)
					return string.Empty;

				InitializeFilter("<stream>", filter);

				return string.Join(" ", GetTexts(filter));
			}
			finally
			{
				if (filter != null)
				{
					Marshal.ReleaseComObject(filter);
					GC.Collect();
					GC.WaitForPendingFinalizers();
				}
			}
		}

		public string GetText(string path)
		{
			IFilter filter = null;

			try
			{
				filter = LoadIFilter(path);

				if (filter == null)
					return string.Empty;

				InitializeFilter(path, filter);

				return string.Join(" ", GetTexts(filter));
			}
			finally
			{
				if (filter != null)
				{
					Marshal.ReleaseComObject(filter);
					GC.Collect();
					GC.WaitForPendingFinalizers();
				}
			}
		}

		private static void InitializeFilter(string path, IFilter filter)
		{
			const FilterSettings iflags = FilterSettings.CanonHyphens
										| FilterSettings.CanonParagraphs
										| FilterSettings.CanonSpaces
										| FilterSettings.ApplyCrawlAttributes
										| FilterSettings.ApplyIndexAttributes
										| FilterSettings.ApplyOtherAttributes
										| FilterSettings.HardLineBreaks
										| FilterSettings.SearchLinks
										| FilterSettings.FilterOwnedValueOk;

			uint i = 0;
			if (filter.Init(iflags, 0, null, ref i) != (int)FilterReturnCodes.Success)
				throw new Exception(string.Format("Could not initialize an IFilter for: '{0}'", path));
		}

		private static IEnumerable<string> GetTexts(IFilter filter)
		{
			StatChunk chunkInfo;

			while (filter.GetChunk(out chunkInfo) == (int)(FilterReturnCodes.Success))
			{
				if (chunkInfo.flags != Chunkstate.ChunkText)
					continue;

				var chunks = GetTextChunks(filter);
				foreach (var chunk in chunks)
					yield return chunk;
			}
		}

		private static IEnumerable<string> GetTextChunks(IFilter filter)
		{
			FilterReturnCodes scode;

			do
			{
				uint pcwcBuffer = 65536;
				var chunkBuffer = new StringBuilder((int)pcwcBuffer);

				scode = (FilterReturnCodes)filter.GetText(ref pcwcBuffer, chunkBuffer);

				if (pcwcBuffer > 0 && chunkBuffer.Length > 0)
				{
					if (chunkBuffer.Length < pcwcBuffer) // Should never happen, but it happens !
						pcwcBuffer = (uint)chunkBuffer.Length;

					yield return chunkBuffer.ToString(0, (int)pcwcBuffer);
				}
			} while (scode == FilterReturnCodes.Success || scode == FilterReturnCodes.LastTextInCurrentChunk);
		}
	}
}

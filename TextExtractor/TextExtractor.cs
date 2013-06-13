using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
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
			if (LoadIFilter(filename, null, ref filter) == (int)IFilterReturnCodes.S_OK)
				return filter;

			return null;
		}

		[DllImport("query.dll", CharSet = CharSet.Unicode)]
		private static extern int LoadIFilter(string pwcsPath, [MarshalAs(UnmanagedType.IUnknown)] object pUnkOuter, ref IFilter ppIUnk);

		[DllImport("query.dll", CharSet = CharSet.Unicode)]
		private static extern int BindIFilterFromStream(System.Runtime.InteropServices.ComTypes.IStream pStm, 
														[MarshalAs(UnmanagedType.IUnknown)] object pUnkOuter,
														[Out] out IFilter ppIUnk);

		[DllImport("ole32.dll", CharSet = CharSet.Unicode)]
		private static extern void CreateStreamOnHGlobal(IntPtr hGlobal, bool fDeleteOnRelease, 
														[Out] out System.Runtime.InteropServices.ComTypes.IStream pStream);

		private static IFilter LoadIFilter(Stream stream)
		{

			// copy stream to byte array
			var b = new byte[stream.Length];
			stream.Read(b, 0, b.Length);

			// allocate space on the native heap
			IntPtr nativePtr = Marshal.AllocHGlobal(b.Length);
	
			// copy byte array to native heap
			Marshal.Copy(b, 0, nativePtr, b.Length);

			// Create a UCOMIStream from the allocated memory
			System.Runtime.InteropServices.ComTypes.IStream comStream;

			CreateStreamOnHGlobal(nativePtr, true, out comStream);

			// Try to load the corresponding IFilter 
			IFilter filter;
			int resultLoad = BindIFilterFromStream(comStream, null, out filter);
			if (resultLoad == (int)IFilterReturnCodes.S_OK)
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
			const IFILTER_INIT iflags = IFILTER_INIT.CANON_HYPHENS |
										IFILTER_INIT.CANON_PARAGRAPHS |
										IFILTER_INIT.CANON_SPACES |
										IFILTER_INIT.APPLY_CRAWL_ATTRIBUTES |
										IFILTER_INIT.APPLY_INDEX_ATTRIBUTES |
										IFILTER_INIT.APPLY_OTHER_ATTRIBUTES |
										IFILTER_INIT.HARD_LINE_BREAKS |
										IFILTER_INIT.SEARCH_LINKS |
										IFILTER_INIT.FILTER_OWNED_VALUE_OK;

			uint i = 0;
			if (filter.Init(iflags, 0, null, ref i) != (int)IFilterReturnCodes.S_OK)
				throw new Exception(string.Format("Could not initialize an IFilter for: '{0}'", path));
		}

		private static IEnumerable<string> GetTexts(IFilter filter)
		{
			STAT_CHUNK chunkInfo;

			while (filter.GetChunk(out chunkInfo) == (int)(IFilterReturnCodes.S_OK))
			{
				if (chunkInfo.flags != CHUNKSTATE.CHUNK_TEXT)
					continue;

				var chunks = GetTextChunks(filter);
				foreach (var chunk in chunks)
					yield return chunk;
			}
		}

		private static IEnumerable<string> GetTextChunks(IFilter filter)
		{
			IFilterReturnCodes scode;

			do
			{
				uint pcwcBuffer = 65536;
				var chunkBuffer = new StringBuilder((int)pcwcBuffer);

				scode = (IFilterReturnCodes)filter.GetText(ref pcwcBuffer, chunkBuffer);

				if (pcwcBuffer > 0 && chunkBuffer.Length > 0)
				{
					if (chunkBuffer.Length < pcwcBuffer) // Should never happen, but it happens !
						pcwcBuffer = (uint)chunkBuffer.Length;

					yield return chunkBuffer.ToString(0, (int)pcwcBuffer);
				}
			} while (scode == IFilterReturnCodes.S_OK || scode == IFilterReturnCodes.FILTER_S_LAST_TEXT);
		}
	}
}

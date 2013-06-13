using System;
using System.Runtime.InteropServices;
using System.Text;

namespace TextExtractor
{
	[Flags]
	public enum IFILTER_INIT : uint
	{
		NONE = 0,
		CANON_PARAGRAPHS = 1,
		HARD_LINE_BREAKS = 2,
		CANON_HYPHENS = 4,
		CANON_SPACES = 8,
		APPLY_INDEX_ATTRIBUTES = 16,
		APPLY_CRAWL_ATTRIBUTES = 256,
		APPLY_OTHER_ATTRIBUTES = 32,
		INDEXING_ONLY = 64,
		SEARCH_LINKS = 128,
		FILTER_OWNED_VALUE_OK = 512
	}

	public enum CHUNK_BREAKTYPE
	{
		CHUNK_NO_BREAK = 0,
		CHUNK_EOW = 1,
		CHUNK_EOS = 2,
		CHUNK_EOP = 3,
		CHUNK_EOC = 4
	}

	[Flags]
	public enum CHUNKSTATE
	{
		CHUNK_TEXT = 0x1,
		CHUNK_VALUE = 0x2,
		CHUNK_FILTER_OWNED_VALUE = 0x4
	}

	[StructLayout(LayoutKind.Sequential)]
	public struct PROPSPEC
	{
		public uint ulKind;
		public uint propid;
		public IntPtr lpwstr;
	}

	[StructLayout(LayoutKind.Sequential)]
	public struct FULLPROPSPEC
	{
		public Guid guidPropSet;
		public PROPSPEC psProperty;
	}

	[StructLayout(LayoutKind.Sequential)]
	public struct STAT_CHUNK
	{
		public uint idChunk;
		[MarshalAs(UnmanagedType.U4)]
		public CHUNK_BREAKTYPE breakType;
		[MarshalAs(UnmanagedType.U4)]
		public CHUNKSTATE flags;
		public uint locale;
		[MarshalAs(UnmanagedType.Struct)]
		public FULLPROPSPEC attribute;
		public uint idChunkSource;
		public uint cwcStartSource;
		public uint cwcLenSource;
	}

	[StructLayout(LayoutKind.Sequential)]
	public struct FILTERREGION
	{
		public uint idChunk;
		public uint cwcStart;
		public uint cwcExtent;
	}

	[ComImport]
	[Guid("89BCB740-6119-101A-BCB7-00DD010655AF")]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	public interface IFilter
	{
		[PreserveSig]
		int Init([MarshalAs(UnmanagedType.U4)] IFILTER_INIT grfFlags, uint cAttributes,
		         [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] FULLPROPSPEC[] aAttributes, ref uint pdwFlags);

		[PreserveSig]
		int GetChunk(out STAT_CHUNK pStat);

		[PreserveSig]
		int GetText(ref uint pcwcBuffer, [MarshalAs(UnmanagedType.LPWStr)] StringBuilder buffer);

		void GetValue(ref UIntPtr ppPropValue);
		void BindRegion([MarshalAs(UnmanagedType.Struct)] FILTERREGION origPos, ref Guid riid, ref UIntPtr ppunk);
	}

	public enum IFilterReturnCodes : uint
	{
		/// Success
		S_OK = 0,

		/// This is the last text in the current chunk
		FILTER_S_LAST_TEXT = 0x00041709,

		/// This is the last value in the current chunk
		FILTER_S_LAST_VALUES = 0x0004170A
	}
}
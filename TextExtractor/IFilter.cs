using System;
using System.Runtime.InteropServices;
using System.Text;

namespace TextExtractor
{
	[Flags]
	public enum FilterSettings : uint
	{
		None = 0,
		CanonParagraphs = 1,
		HardLineBreaks = 2,
		CanonHyphens = 4,
		CanonSpaces = 8,
		ApplyIndexAttributes = 16,
		ApplyCrawlAttributes = 256,
		ApplyOtherAttributes = 32,
		IndexingOnly = 64,
		SearchLinks = 128,
		FilterOwnedValueOk = 512
	}

	public enum ChunkBreaktype
	{
		ChunkNoBreak = 0,
		ChunkEow = 1,
		ChunkEos = 2,
		ChunkEop = 3,
		ChunkEoc = 4
	}

	[Flags]
	public enum Chunkstate
	{
		ChunkText = 0x1,
		ChunkValue = 0x2,
		ChunkFilterOwnedValue = 0x4
	}

	[StructLayout(LayoutKind.Sequential)]
	public struct Propspec
	{
		public uint ulKind;
		public uint propid;
		public IntPtr lpwstr;
	}

	[StructLayout(LayoutKind.Sequential)]
	public struct FullPropSpec
	{
		public Guid guidPropSet;
		public Propspec psProperty;
	}

	[StructLayout(LayoutKind.Sequential)]
	public struct StatChunk
	{
		public uint idChunk;
		[MarshalAs(UnmanagedType.U4)]
		public ChunkBreaktype breakType;
		[MarshalAs(UnmanagedType.U4)]
		public Chunkstate flags;
		public uint locale;
		[MarshalAs(UnmanagedType.Struct)]
		public FullPropSpec attribute;
		public uint idChunkSource;
		public uint cwcStartSource;
		public uint cwcLenSource;
	}

	[StructLayout(LayoutKind.Sequential)]
	public struct FilterRegion
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
		int Init([MarshalAs(UnmanagedType.U4)] FilterSettings grfFlags, uint cAttributes,
		         [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] FullPropSpec[] aAttributes, ref uint pdwFlags);

		[PreserveSig]
		int GetChunk(out StatChunk pStat);

		[PreserveSig]
		int GetText(ref uint pcwcBuffer, [MarshalAs(UnmanagedType.LPWStr)] StringBuilder buffer);

		void GetValue(ref UIntPtr ppPropValue);
		void BindRegion([MarshalAs(UnmanagedType.Struct)] FilterRegion origPos, ref Guid riid, ref UIntPtr ppunk);
	}

	public enum FilterReturnCodes : uint
	{
		/// Success
		Success = 0,

		/// This is the last text in the current chunk
		LastTextInCurrentChunk = 0x00041709,

		/// This is the last value in the current chunk
		LastValueInCurrentChunk = 0x0004170A
	}
}

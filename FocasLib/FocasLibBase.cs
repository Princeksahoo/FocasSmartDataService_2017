using System;
using System.Runtime.InteropServices;
namespace FocasLibrary
{
    public enum CncMachineType
    {
        cncUnknown,
        Series150i,
        Series160i,
        Series180i,
        Series160i180iW,
        Series210i,
        Series0i,
        Series300i,
        Series310i,
        Series320i,
        SeriesPowerMateiD,
        SeriesPowerMateiH
    }
       
    public enum focas_ret
    {
        DNC_CANCEL = -32768,
        DNC_NOFILE = -516,
        DNC_NORMAL = -1,
        DNC_OPENERR = -514,
        DNC_READERR = -517,
        EW_ALARM = 15,
        EW_ATTRIB = 4,
        EW_BUFFER = 10,
        EW_BUS = -11,
        EW_BUSY = -1,
        EW_DATA = 5,
        EW_DTSRVR = 14,
        EW_FUNC = 1,
        EW_HANDLE = -8,
        EW_HSSB = -9,
        EW_LENGTH = 2,
        EW_MMCSYS = -3,
        EW_MODE = 12,
        EW_NODLL = -15,
        EW_NOOPT = 6,
        EW_NOPMC = 1,
        EW_NUMBER = 3,
        EW_OK = 0,
        EW_OVRFLOW = 8,
        EW_PARAM = 9,
        EW_PARITY = -4,
        EW_PASSWD = 0x11,
        EW_PATH = 11,
        EW_PROT = 7,
        EW_PROTOCOL = -17,
        EW_RANGE = 3,
        EW_REJECT = 13,
        EW_RESET = -2,
        EW_SOCKET = -16,
        EW_STOP = 0x10,
        EW_SYSTEM = -5,
        EW_SYSTEM2 = -10,
        EW_TYPE = 4,
        EW_UNEXP = -6,
        EW_VERSION = -7
    }

    public class FocasLibBase
    {
        public const short ALL_AXES = -1;
        public const short ALL_SPINDLES = -1;
        public const short EW_OK = 0;
        public const short MAX_AXIS = 5;
               
        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class _IODBSIG2_data
        {
            public short ent_no;
            public short sig_no;
            public byte sig_name;
            public byte mask_pat;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class _IODBSIG3_data
        {
            public short ent_no;
            public short pmc_no;
            public short sig_no;
            public byte sig_name;
            public byte mask_pat;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ALM_HIS_data
        {
            public short dummy;
            public short alm_grp;
            public short alm_no;
            public byte axis_no;
            public byte year;
            public byte month;
            public byte day;
            public byte hour;
            public byte minute;
            public byte second;
            public byte dummy2;
            public short len_msg;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x20)]
            public string alm_msg = new string(' ', 0x20);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ALM_HIS1
        {
            public FocasLibBase.ALM_HIS_data data1 = new FocasLibBase.ALM_HIS_data();
            public FocasLibBase.ALM_HIS_data data2 = new FocasLibBase.ALM_HIS_data();
            public FocasLibBase.ALM_HIS_data data3 = new FocasLibBase.ALM_HIS_data();
            public FocasLibBase.ALM_HIS_data data4 = new FocasLibBase.ALM_HIS_data();
            public FocasLibBase.ALM_HIS_data data5 = new FocasLibBase.ALM_HIS_data();
            public FocasLibBase.ALM_HIS_data data6 = new FocasLibBase.ALM_HIS_data();
            public FocasLibBase.ALM_HIS_data data7 = new FocasLibBase.ALM_HIS_data();
            public FocasLibBase.ALM_HIS_data data8 = new FocasLibBase.ALM_HIS_data();
            public FocasLibBase.ALM_HIS_data data9 = new FocasLibBase.ALM_HIS_data();
            public FocasLibBase.ALM_HIS_data data10 = new FocasLibBase.ALM_HIS_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ALM_HIS2
        {
            public FocasLibBase.ALM_HIS2_data data1 = new FocasLibBase.ALM_HIS2_data();
            public FocasLibBase.ALM_HIS2_data data2 = new FocasLibBase.ALM_HIS2_data();
            public FocasLibBase.ALM_HIS2_data data3 = new FocasLibBase.ALM_HIS2_data();
            public FocasLibBase.ALM_HIS2_data data4 = new FocasLibBase.ALM_HIS2_data();
            public FocasLibBase.ALM_HIS2_data data5 = new FocasLibBase.ALM_HIS2_data();
            public FocasLibBase.ALM_HIS2_data data6 = new FocasLibBase.ALM_HIS2_data();
            public FocasLibBase.ALM_HIS2_data data7 = new FocasLibBase.ALM_HIS2_data();
            public FocasLibBase.ALM_HIS2_data data8 = new FocasLibBase.ALM_HIS2_data();
            public FocasLibBase.ALM_HIS2_data data9 = new FocasLibBase.ALM_HIS2_data();
            public FocasLibBase.ALM_HIS2_data data10 = new FocasLibBase.ALM_HIS2_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ALM_HIS2_data
        {
            public short alm_grp;
            public short alm_no;
            public short axis_no;
            public short year;
            public short month;
            public short day;
            public short hour;
            public short minute;
            public short second;
            public short len_msg;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x20)]
            public string alm_msg = new string(' ', 0x20);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ALM_HIS3
        {
            public FocasLibBase.ALM_HIS3_data data1 = new FocasLibBase.ALM_HIS3_data();
            public FocasLibBase.ALM_HIS3_data data2 = new FocasLibBase.ALM_HIS3_data();
            public FocasLibBase.ALM_HIS3_data data3 = new FocasLibBase.ALM_HIS3_data();
            public FocasLibBase.ALM_HIS3_data data4 = new FocasLibBase.ALM_HIS3_data();
            public FocasLibBase.ALM_HIS3_data data5 = new FocasLibBase.ALM_HIS3_data();
            public FocasLibBase.ALM_HIS3_data data6 = new FocasLibBase.ALM_HIS3_data();
            public FocasLibBase.ALM_HIS3_data data7 = new FocasLibBase.ALM_HIS3_data();
            public FocasLibBase.ALM_HIS3_data data8 = new FocasLibBase.ALM_HIS3_data();
            public FocasLibBase.ALM_HIS3_data data9 = new FocasLibBase.ALM_HIS3_data();
            public FocasLibBase.ALM_HIS3_data data10 = new FocasLibBase.ALM_HIS3_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ALM_HIS3_data
        {
            public short alm_grp;
            public short alm_no;
            public short axis_no;
            public short year;
            public short month;
            public short day;
            public short hour;
            public short minute;
            public short second;
            public short len_msg;
            public short pth_no;
            public short dummy;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x20)]
            public string alm_msg = new string(' ', 0x20);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ALM_HIS5
        {
            public FocasLibBase.ALM_HIS5_data data1 = new FocasLibBase.ALM_HIS5_data();
            public FocasLibBase.ALM_HIS5_data data2 = new FocasLibBase.ALM_HIS5_data();
            public FocasLibBase.ALM_HIS5_data data3 = new FocasLibBase.ALM_HIS5_data();
            public FocasLibBase.ALM_HIS5_data data4 = new FocasLibBase.ALM_HIS5_data();
            public FocasLibBase.ALM_HIS5_data data5 = new FocasLibBase.ALM_HIS5_data();
            public FocasLibBase.ALM_HIS5_data data6 = new FocasLibBase.ALM_HIS5_data();
            public FocasLibBase.ALM_HIS5_data data7 = new FocasLibBase.ALM_HIS5_data();
            public FocasLibBase.ALM_HIS5_data data8 = new FocasLibBase.ALM_HIS5_data();
            public FocasLibBase.ALM_HIS5_data data9 = new FocasLibBase.ALM_HIS5_data();
            public FocasLibBase.ALM_HIS5_data data10 = new FocasLibBase.ALM_HIS5_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ALM_HIS5_data
        {
            public short alm_grp;
            public short alm_no;
            public short axis_no;
            public short year;
            public short month;
            public short day;
            public short hour;
            public short minute;
            public short second;
            public short len_msg;
            public short pth_no;
            public short dummy;
            public short dsp_flg;
            public short axis_num;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x40)]
            public string alm_msg = new string(' ', 0x40);
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=10)]
            public int[] g_modal;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=10)]
            public byte[] g_dp;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=10)]
            public int[] a_modal;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=10)]
            public byte[] a_dp;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x20)]
            public int[] abs_pos;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x20)]
            public byte[] abs_dp;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x20)]
            public int[] mcn_pos;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x20)]
            public byte[] mcn_dp;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ALMINFO_1
        {
            public FocasLibBase.ALMINFO1_data msg1 = new FocasLibBase.ALMINFO1_data();
            public FocasLibBase.ALMINFO1_data msg2 = new FocasLibBase.ALMINFO1_data();
            public FocasLibBase.ALMINFO1_data msg3 = new FocasLibBase.ALMINFO1_data();
            public FocasLibBase.ALMINFO1_data msg4 = new FocasLibBase.ALMINFO1_data();
            public FocasLibBase.ALMINFO1_data msg5 = new FocasLibBase.ALMINFO1_data();
            public short data_end;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ALMINFO_2
        {
            public FocasLibBase.ALMINFO2_data msg1 = new FocasLibBase.ALMINFO2_data();
            public FocasLibBase.ALMINFO2_data msg2 = new FocasLibBase.ALMINFO2_data();
            public FocasLibBase.ALMINFO2_data msg3 = new FocasLibBase.ALMINFO2_data();
            public FocasLibBase.ALMINFO2_data msg4 = new FocasLibBase.ALMINFO2_data();
            public FocasLibBase.ALMINFO2_data msg5 = new FocasLibBase.ALMINFO2_data();
            public short dataend;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ALMINFO1_data
        {
            public short axis;
            public short alm_no;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ALMINFO2_data
        {
            public short axis;
            public short alm_no;
            public short msg_len;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x20)]
            public string alm_msg = new string(' ', 0x20);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class CFILEINFO
        {
            public FocasLibBase.CFILEINFO_data data1 = new FocasLibBase.CFILEINFO_data();
            public FocasLibBase.CFILEINFO_data data2 = new FocasLibBase.CFILEINFO_data();
            public FocasLibBase.CFILEINFO_data data3 = new FocasLibBase.CFILEINFO_data();
            public FocasLibBase.CFILEINFO_data data4 = new FocasLibBase.CFILEINFO_data();
            public FocasLibBase.CFILEINFO_data data5 = new FocasLibBase.CFILEINFO_data();
            public FocasLibBase.CFILEINFO_data data6 = new FocasLibBase.CFILEINFO_data();
            public FocasLibBase.CFILEINFO_data data7 = new FocasLibBase.CFILEINFO_data();
            public FocasLibBase.CFILEINFO_data data8 = new FocasLibBase.CFILEINFO_data();
            public FocasLibBase.CFILEINFO_data data9 = new FocasLibBase.CFILEINFO_data();
            public FocasLibBase.CFILEINFO_data data10 = new FocasLibBase.CFILEINFO_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class CFILEINFO_data
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=12)]
            public string fname = new string(' ', 12);
            public int file_size;
            public int file_attr;
            public short year;
            public short month;
            public short day;
            public short hour;
            public short minute;
            public short second;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class DIR3_CDATE
        {
            public short year;
            public short month;
            public short day;
            public short hour;
            public short minute;
            public short dummy;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class DIR3_MDATE
        {
            public short year;
            public short month;
            public short day;
            public short hour;
            public short minute;
            public short dummy;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class DIR4_CDATE
        {
            public short year;
            public short month;
            public short day;
            public short hour;
            public short minute;
            public short dummy;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class DIR4_MDATE
        {
            public short year;
            public short month;
            public short day;
            public short hour;
            public short minute;
            public short dummy;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class DSMNTINFO
        {
            public ushort empty_cnt;
            public uint total_size;
            public ushort ReadPtr;
            public ushort WritePtr;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ETBPRM
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=13)]
            public string OwnMACAddress = new string(' ', 0x80);
            public short MaximumChannel;
            public short HDDExistence;
            public short NumberOfScreens;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class FAXIS
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] absolute;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] machine;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] relative;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] distance;
        }        

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class FTPPRM
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x20)]
            public string FTPServerUserName = new string(' ', 0x20);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x20)]
            public string FTPServerPassword = new string(' ', 0x20);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x80)]
            public string FTPServerLoginDirectory = new string(' ', 0x80);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class GASFLOW
        {
            public short slct;
            public short pre_time;
            public short pre_press;
            public short proc_press;
            public short end_time;
            public short end_press;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=3)]
            public short[] reserve = new short[3];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class HOSTPRM
        {
            public short DataServerPort;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x10)]
            public string DataServerIPAddress = new string(' ', 0x10);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x20)]
            public string DataServerUserName = new string(' ', 0x20);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x20)]
            public string DataServerPassword = new string(' ', 0x20);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x80)]
            public string DataServerLoginDirectory = new string(' ', 0x80);
        }

        [StructLayout(LayoutKind.Explicit)]
        public class HSPDATA_1
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8), FieldOffset(0)]
            public byte[] cdatas1 = new byte[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8), FieldOffset(0x120)]
            public byte[] cdatas10 = new byte[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8), FieldOffset(0x20)]
            public byte[] cdatas2 = new byte[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8), FieldOffset(0x40)]
            public byte[] cdatas3 = new byte[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8), FieldOffset(0x60)]
            public byte[] cdatas4 = new byte[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8), FieldOffset(0x80)]
            public byte[] cdatas5 = new byte[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8), FieldOffset(160)]
            public byte[] cdatas6 = new byte[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8), FieldOffset(0xc0)]
            public byte[] cdatas7 = new byte[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8), FieldOffset(0xe0)]
            public byte[] cdatas8 = new byte[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8), FieldOffset(0x100)]
            public byte[] cdatas9 = new byte[8];
        }

        [StructLayout(LayoutKind.Explicit)]
        public class HSPDATA_2
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8), FieldOffset(0)]
            public short[] idatas1 = new short[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8), FieldOffset(0x90)]
            public short[] idatas10 = new short[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8), FieldOffset(0x10)]
            public short[] idatas2 = new short[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8), FieldOffset(0x20)]
            public short[] idatas3 = new short[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8), FieldOffset(0x30)]
            public short[] idatas4 = new short[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8), FieldOffset(0x40)]
            public short[] idatas5 = new short[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8), FieldOffset(80)]
            public short[] idatas6 = new short[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8), FieldOffset(0x60)]
            public short[] idatas7 = new short[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8), FieldOffset(0x70)]
            public short[] idatas8 = new short[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8), FieldOffset(0x80)]
            public short[] idatas9 = new short[8];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class HSPDATA_3
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] ldatas1 = new int[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] ldatas2 = new int[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] ldatas3 = new int[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] ldatas4 = new int[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] ldatas5 = new int[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] ldatas6 = new int[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] ldatas7 = new int[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] ldatas8 = new int[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] ldatas9 = new int[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] ldatas10 = new int[8];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class HSPINFO
        {
            public FocasLibBase.HSPINFO_data prminfo1 = new FocasLibBase.HSPINFO_data();
            public FocasLibBase.HSPINFO_data prminfo2 = new FocasLibBase.HSPINFO_data();
            public FocasLibBase.HSPINFO_data prminfo3 = new FocasLibBase.HSPINFO_data();
            public FocasLibBase.HSPINFO_data prminfo4 = new FocasLibBase.HSPINFO_data();
            public FocasLibBase.HSPINFO_data prminfo5 = new FocasLibBase.HSPINFO_data();
            public FocasLibBase.HSPINFO_data prminfo6 = new FocasLibBase.HSPINFO_data();
            public FocasLibBase.HSPINFO_data prminfo7 = new FocasLibBase.HSPINFO_data();
            public FocasLibBase.HSPINFO_data prminfo8 = new FocasLibBase.HSPINFO_data();
            public FocasLibBase.HSPINFO_data prminfo9 = new FocasLibBase.HSPINFO_data();
            public FocasLibBase.HSPINFO_data prminfo10 = new FocasLibBase.HSPINFO_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class HSPINFO_data
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x10)]
            public byte[] data1 = new byte[0x10];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x10)]
            public byte[] data2 = new byte[0x10];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x10)]
            public byte[] data3 = new byte[0x10];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x10)]
            public byte[] data4 = new byte[0x10];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x10)]
            public byte[] data5 = new byte[0x10];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x10)]
            public byte[] data6 = new byte[0x10];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x10)]
            public byte[] data7 = new byte[0x10];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x10)]
            public byte[] data8 = new byte[0x10];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IDBCHAN
        {
            public FocasLibBase.IDBCHAN_data data1 = new FocasLibBase.IDBCHAN_data();
            public FocasLibBase.IDBCHAN_data data2 = new FocasLibBase.IDBCHAN_data();
            public FocasLibBase.IDBCHAN_data data3 = new FocasLibBase.IDBCHAN_data();
            public FocasLibBase.IDBCHAN_data data4 = new FocasLibBase.IDBCHAN_data();
            public FocasLibBase.IDBCHAN_data data5 = new FocasLibBase.IDBCHAN_data();
            public FocasLibBase.IDBCHAN_data data6 = new FocasLibBase.IDBCHAN_data();
            public FocasLibBase.IDBCHAN_data data7 = new FocasLibBase.IDBCHAN_data();
            public FocasLibBase.IDBCHAN_data data8 = new FocasLibBase.IDBCHAN_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IDBCHAN_data
        {
            public byte chno;
            public sbyte axis;
            public int datanum;
            public ushort datainf;
            public short dataadr;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IDBITD
        {
            public short datano;
            public short type;
            public int data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IDBMGRP
        {
            public short s_no;
            public short dummy;
            public short num;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=500)]
            public short[] group = new short[500];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IDBPDFADIR
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0xd4)]
            public string path = new string(' ', 0xd4);
            public short req_num;
            public short size_kind;
            public short type;
            public short dummy;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IDBPDFSDIR
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0xd4)]
            public string path = new string(' ', 0xd4);
            public short req_num;
            public short dummy;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IDBPDFTDIR
        {
            public uint slct;
            public uint attr;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IDBSFBCHAN
        {
            public byte chno;
            public sbyte axis;
            public ushort shift;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IDBTLM
        {
            public short data_id;
            public FocasLibBase.IDBTLM_item item = new FocasLibBase.IDBTLM_item();
        }

        [StructLayout(LayoutKind.Explicit)]
        public class IDBTLM_item
        {
            [FieldOffset(0)]
            public sbyte data1;
            [FieldOffset(0)]
            public short data2;
            [FieldOffset(0)]
            public int data4;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IDBUNSOLICMSG
        {
            public short getno;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=3)]
            public FocasLibBase.IDBUNSOLICMSG_msg[] msg = new FocasLibBase.IDBUNSOLICMSG_msg[3];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public struct IDBUNSOLICMSG_msg
        {
            public short rdsize;
            public IntPtr data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IDBWRA
        {
            public short datano;
            public short type;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] data = new int[8];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IDBWRC
        {
            public short datano_s;
            public short dummy;
            public short datano_e;
            public FocasLibBase.IDBWRC1 data = new FocasLibBase.IDBWRC1();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IDBWRC_data
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=2)]
            public int[] dummy;
            public int count;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IDBWRC1
        {
            public FocasLibBase.IDBWRC_data data1 = new FocasLibBase.IDBWRC_data();
            public FocasLibBase.IDBWRC_data data2 = new FocasLibBase.IDBWRC_data();
            public FocasLibBase.IDBWRC_data data3 = new FocasLibBase.IDBWRC_data();
            public FocasLibBase.IDBWRC_data data4 = new FocasLibBase.IDBWRC_data();
            public FocasLibBase.IDBWRC_data data5 = new FocasLibBase.IDBWRC_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IDBWRR
        {
            public short datano;
            public short type;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] data = new int[8];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBBTO
        {
            public short datano_s;
            public short type;
            public short datano_e;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x12)]
            public int[] ofs = new int[0x12];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBBUSPRM
        {
            public sbyte fdl_add;
            public sbyte baudrate;
            public ushort tsl;
            public ushort min_tsdr;
            public ushort max_tsdr;
            public byte tqui;
            public byte tset;
            public int ttr;
            public sbyte gap;
            public sbyte hsa;
            public sbyte max_retry;
            public byte bp_flag;
            public ushort min_slv_int;
            public ushort poll_tout;
            public ushort data_cntl;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=6)]
            public byte[] reserve1 = new byte[6];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x20)]
            public byte[] cls2_name = new byte[0x20];
            public short user_dlen;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x3e)]
            public byte[] user_data = new byte[0x3e];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x60)]
            public byte[] reserve2 = new byte[0x60];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBCPRM
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x41)]
            public string NcApli = new string(' ', 0x41);
            public byte Dummy1;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x41)]
            public string HostApli = new string(' ', 0x41);
            public byte Dummy2;
            public uint StatPstv;
            public uint StatNgtv;
            public uint Statmask;
            public uint AlarmStat;
            public uint PsclHaddr;
            public uint PsclLaddr;
            public ushort SvcMode1;
            public ushort SvcMode2;
            public int FileTout;
            public int RemTout;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBDIDO
        {
            public short slave_no;
            public short slot_no;
            public byte di_size;
            public char di_type;
            public ushort di_addr;
            public byte do_size;
            public char do_type;
            public ushort do_addr;
            public short shift;
            public byte module_dlen;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x80)]
            public char[] module_data = new char[0x80];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBDSSET
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x10)]
            public string host_ip = new string(' ', 0x10);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x20)]
            public string host_uname = new string(' ', 0x20);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x20)]
            public string host_passwd = new string(' ', 0x20);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x80)]
            public string host_dir = new string(' ', 0x80);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=13)]
            public string dtsv_mac = new string(' ', 13);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x10)]
            public string dtsv_ip = new string(' ', 0x10);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x10)]
            public string dtsv_mask = new string(' ', 0x10);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBEDGE
        {
            public FocasLibBase.IODBEDGE_data data1 = new FocasLibBase.IODBEDGE_data();
            public FocasLibBase.IODBEDGE_data data2 = new FocasLibBase.IODBEDGE_data();
            public FocasLibBase.IODBEDGE_data data3 = new FocasLibBase.IODBEDGE_data();
            public FocasLibBase.IODBEDGE_data data4 = new FocasLibBase.IODBEDGE_data();
            public FocasLibBase.IODBEDGE_data data5 = new FocasLibBase.IODBEDGE_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBEDGE_data
        {
            public short slct;
            public short angle;
            public short power;
            public short freq;
            public short duty;
            public int pier_t;
            public short g_press;
            public short g_kind;
            public int r_len;
            public short r_feed;
            public short r_freq;
            public short r_duty;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=5)]
            public short[] reserve = new short[5];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBETP
        {
            public short Dummy_ParameterType;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=210)]
            public byte[] prm = new byte[210];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBETP_ETB
        {
            public short ParameterType;
            public FocasLibBase.ETBPRM etb;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBETP_FTP
        {
            public short ParameterType;
            public FocasLibBase.FTPPRM ftp;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBETP_HOST
        {
            public short ParameterType;
            public FocasLibBase.HOSTPRM host;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBETP_TCP
        {
            public short ParameterType;
            public FocasLibBase.TCPPRM tcp;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBGNRL
        {
            public short datano;
            public short type;
            public byte sgnal;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBHPAC
        {
            public FocasLibBase.IODBHPAC_tune tune1 = new FocasLibBase.IODBHPAC_tune();
            public FocasLibBase.IODBHPAC_tune tune2 = new FocasLibBase.IODBHPAC_tune();
            public FocasLibBase.IODBHPAC_tune tune3 = new FocasLibBase.IODBHPAC_tune();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBHPAC_tune
        {
            public short slct;
            public short diff;
            public short fine;
            public short acc_lv;
            public int bipl;
            public short aipl;
            public int corner;
            public int clamp;
            public int c_acc;
            public int foward;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] reserve = new int[8];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBHPPR
        {
            public FocasLibBase.IODBHPPR_tune tune1 = new FocasLibBase.IODBHPPR_tune();
            public FocasLibBase.IODBHPPR_tune tune2 = new FocasLibBase.IODBHPPR_tune();
            public FocasLibBase.IODBHPPR_tune tune3 = new FocasLibBase.IODBHPPR_tune();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBHPPR_tune
        {
            public short slct;
            public short diff;
            public short fine;
            public short acc_lv;
            public int max_f;
            public short bipl;
            public short aipl;
            public int corner;
            public short clamp;
            public int radius;
            public int max_cf;
            public int min_cf;
            public int foward;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=5)]
            public int[] reserve = new int[5];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBHPST
        {
            public short slct;
            public short hpcc;
            public short multi;
            public short ovr1;
            public short ign_f;
            public short foward;
            public int max_f;
            public short ovr2;
            public short ovr3;
            public short ovr4;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=7)]
            public int[] reserve = new int[7];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBIDINF
        {
            public int id_no;
            public short drv_no;
            public short acc_element;
            public short err_general;
            public short err_id_no;
            public short err_id_name;
            public short err_attr;
            public short err_unit;
            public short err_min_val;
            public short err_max_val;
            public short id_name_len;
            public short id_name_max;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=60)]
            public string id_name = new string(' ', 60);
            public int attr;
            public short unit_len;
            public short unit_max;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=12)]
            public byte[] unit = new byte[12];
            public int min_val;
            public int max_val;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBINDEADR
        {
            public byte dummy;
            public char indi_type;
            public ushort indi_addr;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBINT
        {
            public short datano_s;
            public short type;
            public short datano_e;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x18)]
            public int[] data = new int[0x18];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBITEM
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x3e)]
            public string name1 = new string(' ', 0x3e);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x3e)]
            public string name2 = new string(' ', 0x3e);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x3e)]
            public string name3 = new string(' ', 0x3e);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x3e)]
            public string name4 = new string(' ', 0x3e);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x3e)]
            public string name5 = new string(' ', 0x3e);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x3e)]
            public string name6 = new string(' ', 0x3e);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x3e)]
            public string name7 = new string(' ', 0x3e);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x3e)]
            public string name8 = new string(' ', 0x3e);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x3e)]
            public string name9 = new string(' ', 0x3e);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x3e)]
            public string name10 = new string(' ', 0x3e);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBLAGSL
        {
            public short slct;
            public short ag_slt;
            public short agflow_slt;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=6)]
            public short[] reserve = new short[6];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBLAGST
        {
            public FocasLibBase.GASFLOW data1 = new FocasLibBase.GASFLOW();
            public FocasLibBase.GASFLOW data2 = new FocasLibBase.GASFLOW();
            public FocasLibBase.GASFLOW data3 = new FocasLibBase.GASFLOW();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBLEGPR
        {
            public short slct;
            public short power;
            public short freq;
            public short duty;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=5)]
            public short[] reserve = new short[5];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBLPCPR
        {
            public short slct;
            public short power;
            public short freq;
            public short duty;
            public int time;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=4)]
            public short[] reserve = new short[4];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBLPWDT
        {
            public short slct;
            public short dty_const;
            public short dty_min;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=6)]
            public short[] reserve = new short[6];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBMLTTL
        {
            public FocasLibBase.IODBMLTTL_data data1 = new FocasLibBase.IODBMLTTL_data();
            public FocasLibBase.IODBMLTTL_data data2 = new FocasLibBase.IODBMLTTL_data();
            public FocasLibBase.IODBMLTTL_data data3 = new FocasLibBase.IODBMLTTL_data();
            public FocasLibBase.IODBMLTTL_data data4 = new FocasLibBase.IODBMLTTL_data();
            public FocasLibBase.IODBMLTTL_data data5 = new FocasLibBase.IODBMLTTL_data();
            public FocasLibBase.IODBMLTTL_data data6 = new FocasLibBase.IODBMLTTL_data();
            public FocasLibBase.IODBMLTTL_data data7 = new FocasLibBase.IODBMLTTL_data();
            public FocasLibBase.IODBMLTTL_data data8 = new FocasLibBase.IODBMLTTL_data();
            public FocasLibBase.IODBMLTTL_data data9 = new FocasLibBase.IODBMLTTL_data();
            public FocasLibBase.IODBMLTTL_data data10 = new FocasLibBase.IODBMLTTL_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBMLTTL_data
        {
            public short slct;
            public short m_tl_no;
            public int m_tl_radius;
            public int m_tl_angle;
            public int x_axis_ofs;
            public int y_axis_ofs;
            public byte tl_shape;
            public int tl_size_i;
            public int tl_size_j;
            public int tl_angle;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=7)]
            public int[] reserve = new int[7];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBMNGTIME
        {
            public FocasLibBase.IODBMNGTIME_data data1 = new FocasLibBase.IODBMNGTIME_data();
            public FocasLibBase.IODBMNGTIME_data data2 = new FocasLibBase.IODBMNGTIME_data();
            public FocasLibBase.IODBMNGTIME_data data3 = new FocasLibBase.IODBMNGTIME_data();
            public FocasLibBase.IODBMNGTIME_data data4 = new FocasLibBase.IODBMNGTIME_data();
            public FocasLibBase.IODBMNGTIME_data data5 = new FocasLibBase.IODBMNGTIME_data();
            public FocasLibBase.IODBMNGTIME_data data6 = new FocasLibBase.IODBMNGTIME_data();
            public FocasLibBase.IODBMNGTIME_data data7 = new FocasLibBase.IODBMNGTIME_data();
            public FocasLibBase.IODBMNGTIME_data data8 = new FocasLibBase.IODBMNGTIME_data();
            public FocasLibBase.IODBMNGTIME_data data9 = new FocasLibBase.IODBMNGTIME_data();
            public FocasLibBase.IODBMNGTIME_data data10 = new FocasLibBase.IODBMNGTIME_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBMNGTIME_data
        {
            public uint life;
            public uint total;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBMR
        {
            public short datano_s;
            public short dummy;
            public short datano_e;
            public FocasLibBase.IODBMR1 data = new FocasLibBase.IODBMR1();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBMR_data
        {
            public int mcr_val;
            public short dec_val;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBMR1
        {
            public FocasLibBase.IODBMR_data data1 = new FocasLibBase.IODBMR_data();
            public FocasLibBase.IODBMR_data data2 = new FocasLibBase.IODBMR_data();
            public FocasLibBase.IODBMR_data data3 = new FocasLibBase.IODBMR_data();
            public FocasLibBase.IODBMR_data data4 = new FocasLibBase.IODBMR_data();
            public FocasLibBase.IODBMR_data data5 = new FocasLibBase.IODBMR_data();
            public FocasLibBase.IODBMR_data data6 = new FocasLibBase.IODBMR_data();
            public FocasLibBase.IODBMR_data data7 = new FocasLibBase.IODBMR_data();
            public FocasLibBase.IODBMR_data data8 = new FocasLibBase.IODBMR_data();
            public FocasLibBase.IODBMR_data data9 = new FocasLibBase.IODBMR_data();
            public FocasLibBase.IODBMR_data data10 = new FocasLibBase.IODBMR_data();
            public FocasLibBase.IODBMR_data data11 = new FocasLibBase.IODBMR_data();
            public FocasLibBase.IODBMR_data data12 = new FocasLibBase.IODBMR_data();
            public FocasLibBase.IODBMR_data data13 = new FocasLibBase.IODBMR_data();
            public FocasLibBase.IODBMR_data data14 = new FocasLibBase.IODBMR_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBMSTP
        {
            public short datano_s;
            public short dummy;
            public short datano_e;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=7)]
            public sbyte[] data = new sbyte[7];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBMTAP
        {
            public FocasLibBase.IODBMTAP_data data1 = new FocasLibBase.IODBMTAP_data();
            public FocasLibBase.IODBMTAP_data data2 = new FocasLibBase.IODBMTAP_data();
            public FocasLibBase.IODBMTAP_data data3 = new FocasLibBase.IODBMTAP_data();
            public FocasLibBase.IODBMTAP_data data4 = new FocasLibBase.IODBMTAP_data();
            public FocasLibBase.IODBMTAP_data data5 = new FocasLibBase.IODBMTAP_data();
            public FocasLibBase.IODBMTAP_data data6 = new FocasLibBase.IODBMTAP_data();
            public FocasLibBase.IODBMTAP_data data7 = new FocasLibBase.IODBMTAP_data();
            public FocasLibBase.IODBMTAP_data data8 = new FocasLibBase.IODBMTAP_data();
            public FocasLibBase.IODBMTAP_data data9 = new FocasLibBase.IODBMTAP_data();
            public FocasLibBase.IODBMTAP_data data10 = new FocasLibBase.IODBMTAP_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBMTAP_data
        {
            public short slct;
            public int tool_no;
            public int x_axis_ofs;
            public int y_axis_ofs;
            public int punch_count;
            public int tool_life;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=11)]
            public int[] reserve = new int[11];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBOVL
        {
            public short datano;
            public short type;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x10)]
            public int[] data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBPI
        {
            public short datano_s;
            public short dummy;
            public short datano_e;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=5)]
            public sbyte[] data = new sbyte[5];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBPIRC
        {
            public FocasLibBase.IODBPIRC_data data1 = new FocasLibBase.IODBPIRC_data();
            public FocasLibBase.IODBPIRC_data data2 = new FocasLibBase.IODBPIRC_data();
            public FocasLibBase.IODBPIRC_data data3 = new FocasLibBase.IODBPIRC_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBPIRC_data
        {
            public short slct;
            public short power;
            public short freq;
            public short duty;
            public short i_freq;
            public short i_duty;
            public short step_t;
            public short step_sum;
            public int pier_t;
            public short g_press;
            public short g_kind;
            public short g_time;
            public short def_pos;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=4)]
            public short[] reserve = new short[4];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBPMAINTE
        {
            public FocasLibBase.IODBPMAINTE_data data1 = new FocasLibBase.IODBPMAINTE_data();
            public FocasLibBase.IODBPMAINTE_data data2 = new FocasLibBase.IODBPMAINTE_data();
            public FocasLibBase.IODBPMAINTE_data data3 = new FocasLibBase.IODBPMAINTE_data();
            public FocasLibBase.IODBPMAINTE_data data4 = new FocasLibBase.IODBPMAINTE_data();
            public FocasLibBase.IODBPMAINTE_data data5 = new FocasLibBase.IODBPMAINTE_data();
            public FocasLibBase.IODBPMAINTE_data data6 = new FocasLibBase.IODBPMAINTE_data();
            public FocasLibBase.IODBPMAINTE_data data7 = new FocasLibBase.IODBPMAINTE_data();
            public FocasLibBase.IODBPMAINTE_data data8 = new FocasLibBase.IODBPMAINTE_data();
            public FocasLibBase.IODBPMAINTE_data data9 = new FocasLibBase.IODBPMAINTE_data();
            public FocasLibBase.IODBPMAINTE_data data10 = new FocasLibBase.IODBPMAINTE_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBPMAINTE_data
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x3e)]
            public string name = new string(' ', 0x3e);
            public int type;
            public int total;
            public int remain;
            public int stat;
        }

        [StructLayout(LayoutKind.Explicit)]
        public class IODBPMC0
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=10), FieldOffset(8)]
            public byte[] cdata;
            [FieldOffset(6)]
            public short datano_e;
            [FieldOffset(4)]
            public short datano_s;
            [FieldOffset(0)]
            public short type_a;
            [FieldOffset(2)]
            public short type_d;
        }

        [StructLayout(LayoutKind.Explicit)]
        public class IODBPMC1
        {
            [FieldOffset(6)]
            public short datano_e;
            [FieldOffset(4)]
            public short datano_s;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=10), FieldOffset(8)]
            public short[] idata;
            [FieldOffset(0)]
            public short type_a;
            [FieldOffset(2)]
            public short type_d;
        }

        [StructLayout(LayoutKind.Explicit)]
        public class IODBPMC2
        {
            [FieldOffset(6)]
            public short datano_e;
            [FieldOffset(4)]
            public short datano_s;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=10), FieldOffset(8)]
            public int[] ldata;
            [FieldOffset(0)]
            public short type_a;
            [FieldOffset(2)]
            public short type_d;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBPMCCNTL
        {
            public short datano_s;
            public short dummy;
            public short datano_e;
            public FocasLibBase.IODBPMCCNTL1 info = new FocasLibBase.IODBPMCCNTL1();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBPMCCNTL_info
        {
            public byte tbl_prm;
            public byte data_type;
            public ushort data_size;
            public ushort data_dsp;
            public short dummy;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBPMCCNTL1
        {
            public FocasLibBase.IODBPMCCNTL_info info1 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info2 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info3 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info4 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info5 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info6 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info7 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info8 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info9 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info10 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info11 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info12 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info13 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info14 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info15 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info16 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info17 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info18 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info19 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info20 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info21 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info22 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info23 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info24 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info25 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info26 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info27 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info28 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info29 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info30 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info31 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info32 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info33 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info34 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info35 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info36 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info37 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info38 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info39 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info40 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info41 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info42 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info43 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info44 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info45 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info46 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info47 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info48 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info49 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info50 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info51 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info52 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info53 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info54 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info55 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info56 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info57 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info58 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info59 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info60 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info61 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info62 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info63 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info64 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info65 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info66 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info67 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info68 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info69 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info70 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info71 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info72 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info73 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info74 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info75 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info76 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info77 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info78 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info79 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info80 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info81 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info82 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info83 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info84 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info85 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info86 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info87 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info88 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info89 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info90 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info91 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info92 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info93 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info94 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info95 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info96 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info97 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info98 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info99 = new FocasLibBase.IODBPMCCNTL_info();
            public FocasLibBase.IODBPMCCNTL_info info100 = new FocasLibBase.IODBPMCCNTL_info();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBPMCEXT
        {
            public short type_a;
            public short type_d;
            public short datano_s;
            public short datano_e;
            public short err_code;
            public short reserved;
            [MarshalAs(UnmanagedType.AsAny)]
            public object data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBPR
        {
            public int datano_s;
            public short dummy;
            public int datano_e;
            public FocasLibBase.IODBPR1 data = new FocasLibBase.IODBPR1();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBPR_data
        {
            public int mcr_val;
            public short dec_val;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBPR1
        {
            public FocasLibBase.IODBPR_data data1 = new FocasLibBase.IODBPR_data();
            public FocasLibBase.IODBPR_data data2 = new FocasLibBase.IODBPR_data();
            public FocasLibBase.IODBPR_data data3 = new FocasLibBase.IODBPR_data();
            public FocasLibBase.IODBPR_data data4 = new FocasLibBase.IODBPR_data();
            public FocasLibBase.IODBPR_data data5 = new FocasLibBase.IODBPR_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBPRFADR
        {
            public byte di_size;
            public byte di_type;
            public ushort di_addr;
            public short reserve1;
            public byte do_size;
            public byte do_type;
            public ushort do_addr;
            public short reserve2;
            public byte dgn_size;
            public byte dgn_type;
            public ushort dgn_addr;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBPRM
        {
            public FocasLibBase.IODBPRM2 prm1 = new FocasLibBase.IODBPRM2();
            public FocasLibBase.IODBPRM2 prm2 = new FocasLibBase.IODBPRM2();
            public FocasLibBase.IODBPRM2 prm3 = new FocasLibBase.IODBPRM2();
            public FocasLibBase.IODBPRM2 prm4 = new FocasLibBase.IODBPRM2();
            public FocasLibBase.IODBPRM2 prm5 = new FocasLibBase.IODBPRM2();
            public FocasLibBase.IODBPRM2 prm6 = new FocasLibBase.IODBPRM2();
            public FocasLibBase.IODBPRM2 prm7 = new FocasLibBase.IODBPRM2();
            public FocasLibBase.IODBPRM2 prm8 = new FocasLibBase.IODBPRM2();
            public FocasLibBase.IODBPRM2 prm9 = new FocasLibBase.IODBPRM2();
            public FocasLibBase.IODBPRM2 prm10 = new FocasLibBase.IODBPRM2();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBPRM_data
        {
            public int prm_val;
            public int dec_val;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBPRM1
        {
            public FocasLibBase.IODBPRM_data data1 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data2 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data3 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data4 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data5 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data6 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data7 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data8 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data9 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data10 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data11 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data12 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data13 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data14 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data15 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data16 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data17 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data18 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data19 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data20 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data21 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data22 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data23 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data24 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data25 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data26 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data27 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data28 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data29 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data30 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data31 = new FocasLibBase.IODBPRM_data();
            public FocasLibBase.IODBPRM_data data32 = new FocasLibBase.IODBPRM_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBPRM2
        {
            public int datano;
            public short type;
            public short axis;
            public short info;
            public short unit;
            public FocasLibBase.IODBPRM1 data = new FocasLibBase.IODBPRM1();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBPRMNO
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=10)]
            public int[] prm = new int[10];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBPSCD
        {
            public FocasLibBase.IODBPSCD_data data1 = new FocasLibBase.IODBPSCD_data();
            public FocasLibBase.IODBPSCD_data data2 = new FocasLibBase.IODBPSCD_data();
            public FocasLibBase.IODBPSCD_data data3 = new FocasLibBase.IODBPSCD_data();
            public FocasLibBase.IODBPSCD_data data4 = new FocasLibBase.IODBPSCD_data();
            public FocasLibBase.IODBPSCD_data data5 = new FocasLibBase.IODBPSCD_data();
            public FocasLibBase.IODBPSCD_data data6 = new FocasLibBase.IODBPSCD_data();
            public FocasLibBase.IODBPSCD_data data7 = new FocasLibBase.IODBPSCD_data();
            public FocasLibBase.IODBPSCD_data data8 = new FocasLibBase.IODBPSCD_data();
            public FocasLibBase.IODBPSCD_data data9 = new FocasLibBase.IODBPSCD_data();
            public FocasLibBase.IODBPSCD_data data10 = new FocasLibBase.IODBPSCD_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBPSCD_data
        {
            public short slct;
            public int feed;
            public short power;
            public short freq;
            public short duty;
            public short g_press;
            public short g_kind;
            public short g_ready_t;
            public short displace;
            public int supple;
            public short edge_slt;
            public short appr_slt;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=5)]
            public short[] reserve = new short[5];
        }

        [StructLayout(LayoutKind.Explicit)]
        public class IODBPSD_1
        {
            [FieldOffset(4)]
            public byte cdata;
            [FieldOffset(0)]
            public short datano;
            [FieldOffset(4)]
            public short idata;
            [FieldOffset(4)]
            public int ldata;
            [FieldOffset(2)]
            public short type;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBPSD_2
        {
            public short datano;
            public short type;
            public FocasLibBase.REALPRM rdata = new FocasLibBase.REALPRM();
        }

        [StructLayout(LayoutKind.Explicit)]
        public class IODBPSD_3
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8), FieldOffset(4)]
            public byte[] cdatas = new byte[8];
            [FieldOffset(0)]
            public short datano;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8), FieldOffset(4)]
            public short[] idatas = new short[8];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8), FieldOffset(4)]
            public int[] ldatas = new int[8];
            [FieldOffset(2)]
            public short type;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBPSD_4
        {
            public short datano;
            public short type;
            public FocasLibBase.REALPRMS rdatas = new FocasLibBase.REALPRMS();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBPSD_A
        {
            public FocasLibBase.IODBPSD_1 data1 = new FocasLibBase.IODBPSD_1();
            public FocasLibBase.IODBPSD_1 data2 = new FocasLibBase.IODBPSD_1();
            public FocasLibBase.IODBPSD_1 data3 = new FocasLibBase.IODBPSD_1();
            public FocasLibBase.IODBPSD_1 data4 = new FocasLibBase.IODBPSD_1();
            public FocasLibBase.IODBPSD_1 data5 = new FocasLibBase.IODBPSD_1();
            public FocasLibBase.IODBPSD_1 data6 = new FocasLibBase.IODBPSD_1();
            public FocasLibBase.IODBPSD_1 data7 = new FocasLibBase.IODBPSD_1();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBPSD_B
        {
            public FocasLibBase.IODBPSD_2 data1 = new FocasLibBase.IODBPSD_2();
            public FocasLibBase.IODBPSD_2 data2 = new FocasLibBase.IODBPSD_2();
            public FocasLibBase.IODBPSD_2 data3 = new FocasLibBase.IODBPSD_2();
            public FocasLibBase.IODBPSD_2 data4 = new FocasLibBase.IODBPSD_2();
            public FocasLibBase.IODBPSD_2 data5 = new FocasLibBase.IODBPSD_2();
            public FocasLibBase.IODBPSD_2 data6 = new FocasLibBase.IODBPSD_2();
            public FocasLibBase.IODBPSD_2 data7 = new FocasLibBase.IODBPSD_2();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBPSD_C
        {
            public FocasLibBase.IODBPSD_3 data1 = new FocasLibBase.IODBPSD_3();
            public FocasLibBase.IODBPSD_3 data2 = new FocasLibBase.IODBPSD_3();
            public FocasLibBase.IODBPSD_3 data3 = new FocasLibBase.IODBPSD_3();
            public FocasLibBase.IODBPSD_3 data4 = new FocasLibBase.IODBPSD_3();
            public FocasLibBase.IODBPSD_3 data5 = new FocasLibBase.IODBPSD_3();
            public FocasLibBase.IODBPSD_3 data6 = new FocasLibBase.IODBPSD_3();
            public FocasLibBase.IODBPSD_3 data7 = new FocasLibBase.IODBPSD_3();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBPSD_D
        {
            public FocasLibBase.IODBPSD_4 data1 = new FocasLibBase.IODBPSD_4();
            public FocasLibBase.IODBPSD_4 data2 = new FocasLibBase.IODBPSD_4();
            public FocasLibBase.IODBPSD_4 data3 = new FocasLibBase.IODBPSD_4();
            public FocasLibBase.IODBPSD_4 data4 = new FocasLibBase.IODBPSD_4();
            public FocasLibBase.IODBPSD_4 data5 = new FocasLibBase.IODBPSD_4();
            public FocasLibBase.IODBPSD_4 data6 = new FocasLibBase.IODBPSD_4();
            public FocasLibBase.IODBPSD_4 data7 = new FocasLibBase.IODBPSD_4();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBRDNA
        {
            public short datano;
            public short type;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=9)]
            public string sgnl1_name = new string(' ', 9);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=9)]
            public string sgnl2_name = new string(' ', 9);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=9)]
            public string sgnl3_name = new string(' ', 9);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=9)]
            public string sgnl4_name = new string(' ', 9);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=9)]
            public string sgnl5_name = new string(' ', 9);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=9)]
            public string sgnl6_name = new string(' ', 9);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=9)]
            public string sgnl7_name = new string(' ', 9);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=9)]
            public string sgnl8_name = new string(' ', 9);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBRMTPRM
        {
            public short condition;
            public short reserve;
            public FocasLibBase.IODBRMTPRM_trg trg = new FocasLibBase.IODBRMTPRM_trg();
            public int delay;
            public short wv_intrvl;
            public short io_intrvl;
            public short kind1;
            public short kind2;
            public FocasLibBase.IODBRMTPRM1 ampl = new FocasLibBase.IODBRMTPRM1();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBRMTPRM_alm
        {
            public short no;
            public sbyte axis;
            public byte type;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBRMTPRM_io
        {
            public char adr;
            public byte bit;
            public short no;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBRMTPRM_smpl
        {
            public char adr;
            public byte bit;
            public short no;
        }

        [StructLayout(LayoutKind.Explicit)]
        public class IODBRMTPRM_trg
        {
            [FieldOffset(0)]
            public FocasLibBase.IODBRMTPRM_alm alm = new FocasLibBase.IODBRMTPRM_alm();
            [FieldOffset(0)]
            public FocasLibBase.IODBRMTPRM_io io = new FocasLibBase.IODBRMTPRM_io();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBRMTPRM1
        {
            public FocasLibBase.IODBRMTPRM_smpl ampl1 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl2 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl3 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl4 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl5 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl6 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl7 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl8 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl9 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl10 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl11 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl12 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl13 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl14 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl15 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl16 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl17 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl18 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl19 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl20 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl21 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl22 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl23 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl24 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl25 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl26 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl27 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl28 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl29 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl30 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl31 = new FocasLibBase.IODBRMTPRM_smpl();
            public FocasLibBase.IODBRMTPRM_smpl ampl32 = new FocasLibBase.IODBRMTPRM_smpl();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBSAFE
        {
            public FocasLibBase.IODBSAFE_data data1 = new FocasLibBase.IODBSAFE_data();
            public FocasLibBase.IODBSAFE_data data2 = new FocasLibBase.IODBSAFE_data();
            public FocasLibBase.IODBSAFE_data data3 = new FocasLibBase.IODBSAFE_data();
            public FocasLibBase.IODBSAFE_data data4 = new FocasLibBase.IODBSAFE_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBSAFE_data
        {
            public short slct;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=3)]
            public int[] data = new int[3];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBSGNL
        {
            public short datano;
            public short type;
            public short mode;
            public short hndl_ax;
            public short hndl_mv;
            public short rpd_ovrd;
            public short jog_ovrd;
            public short feed_ovrd;
            public short spdl_ovrd;
            public short blck_del;
            public short sngl_blck;
            public short machn_lock;
            public short dry_run;
            public short mem_prtct;
            public short feed_hold;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBSIG
        {
            public short datano;
            public short type;
            public FocasLibBase.IODBSIG1 data = new FocasLibBase.IODBSIG1();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBSIG_data
        {
            public short ent_no;
            public short sig_no;
            public byte sig_name;
            public byte mask_pat;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBSIG1
        {
            public FocasLibBase.IODBSIG_data data1 = new FocasLibBase.IODBSIG_data();
            public FocasLibBase.IODBSIG_data data2 = new FocasLibBase.IODBSIG_data();
            public FocasLibBase.IODBSIG_data data3 = new FocasLibBase.IODBSIG_data();
            public FocasLibBase.IODBSIG_data data4 = new FocasLibBase.IODBSIG_data();
            public FocasLibBase.IODBSIG_data data5 = new FocasLibBase.IODBSIG_data();
            public FocasLibBase.IODBSIG_data data6 = new FocasLibBase.IODBSIG_data();
            public FocasLibBase.IODBSIG_data data7 = new FocasLibBase.IODBSIG_data();
            public FocasLibBase.IODBSIG_data data8 = new FocasLibBase.IODBSIG_data();
            public FocasLibBase.IODBSIG_data data9 = new FocasLibBase.IODBSIG_data();
            public FocasLibBase.IODBSIG_data data10 = new FocasLibBase.IODBSIG_data();
            public FocasLibBase.IODBSIG_data data11 = new FocasLibBase.IODBSIG_data();
            public FocasLibBase.IODBSIG_data data12 = new FocasLibBase.IODBSIG_data();
            public FocasLibBase.IODBSIG_data data13 = new FocasLibBase.IODBSIG_data();
            public FocasLibBase.IODBSIG_data data14 = new FocasLibBase.IODBSIG_data();
            public FocasLibBase.IODBSIG_data data15 = new FocasLibBase.IODBSIG_data();
            public FocasLibBase.IODBSIG_data data16 = new FocasLibBase.IODBSIG_data();
            public FocasLibBase.IODBSIG_data data17 = new FocasLibBase.IODBSIG_data();
            public FocasLibBase.IODBSIG_data data18 = new FocasLibBase.IODBSIG_data();
            public FocasLibBase.IODBSIG_data data19 = new FocasLibBase.IODBSIG_data();
            public FocasLibBase.IODBSIG_data data20 = new FocasLibBase.IODBSIG_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBSIG2
        {
            public short datano;
            public short type;
            public FocasLibBase.IODBSIG2_data data = new FocasLibBase.IODBSIG2_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBSIG2_data
        {
            public FocasLibBase._IODBSIG2_data data1 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data2 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data3 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data4 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data5 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data6 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data7 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data8 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data9 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data10 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data11 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data12 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data13 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data14 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data15 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data16 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data17 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data18 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data19 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data20 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data31 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data32 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data33 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data34 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data35 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data36 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data37 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data38 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data39 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data40 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data41 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data42 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data43 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data44 = new FocasLibBase._IODBSIG2_data();
            public FocasLibBase._IODBSIG2_data data45 = new FocasLibBase._IODBSIG2_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBSIG3
        {
            public short datano;
            public short type;
            public FocasLibBase.IODBSIG3_data data = new FocasLibBase.IODBSIG3_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBSIG3_data
        {
            public FocasLibBase._IODBSIG3_data data1 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data2 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data3 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data4 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data5 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data6 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data7 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data8 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data9 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data10 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data11 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data12 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data13 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data14 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data15 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data16 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data17 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data18 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data19 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data20 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data21 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data22 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data23 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data24 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data25 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data26 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data27 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data28 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data29 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data30 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data31 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data32 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data33 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data34 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data35 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data36 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data37 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data38 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data39 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data40 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data41 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data42 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data43 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data44 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data45 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data46 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data47 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data48 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data49 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data50 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data51 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data52 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data53 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data54 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data55 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data56 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data57 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data58 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data59 = new FocasLibBase._IODBSIG3_data();
            public FocasLibBase._IODBSIG3_data data60 = new FocasLibBase._IODBSIG3_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBSIGAD
        {
            public byte adr;
            public byte reserve;
            public short no;
            public short size;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBSLOP
        {
            public FocasLibBase.IODBSLOP_data data1 = new FocasLibBase.IODBSLOP_data();
            public FocasLibBase.IODBSLOP_data data2 = new FocasLibBase.IODBSLOP_data();
            public FocasLibBase.IODBSLOP_data data3 = new FocasLibBase.IODBSLOP_data();
            public FocasLibBase.IODBSLOP_data data4 = new FocasLibBase.IODBSLOP_data();
            public FocasLibBase.IODBSLOP_data data5 = new FocasLibBase.IODBSLOP_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBSLOP_data
        {
            public int slct;
            public int upleng;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=10)]
            public short[] upsp = new short[10];
            public int dwleng;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=10)]
            public short[] dwsp = new short[10];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=10)]
            public short[] reserve = new short[10];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBSLVADR
        {
            public byte slave_no;
            public byte di_size;
            public byte di_type;
            public ushort di_addr;
            public byte do_size;
            public byte do_type;
            public ushort do_addr;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=7)]
            public byte[] reserve = new byte[7];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBSLVID
        {
            public short dis_enb;
            public short slave_no;
            public short nsl;
            public byte dgn_size;
            public char dgn_type;
            public ushort dgn_addr;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBSLVPRM
        {
            public short dis_enb;
            public ushort ident_no;
            public byte slv_flag;
            public byte slv_type;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=12)]
            public byte[] reserve1 = new byte[12];
            public byte slv_stat;
            public byte wd_fact1;
            public byte wd_fact2;
            public byte min_tsdr;
            public char reserve2;
            public byte grp_ident;
            public short user_plen;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x20)]
            public byte[] user_pdata = new byte[0x20];
            public short cnfg_dlen;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x7e)]
            public byte[] cnfg_data = new byte[0x7e];
            public short slv_ulen;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=30)]
            public byte[] slv_udata = new byte[30];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public byte[] reserve3 = new byte[8];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBSLVPRM2
        {
            public short dis_enb;
            public ushort ident_no;
            public byte slv_flag;
            public byte slv_type;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=12)]
            public byte[] reserve1 = new byte[12];
            public byte slv_stat;
            public byte wd_fact1;
            public byte wd_fact2;
            public byte min_tsdr;
            public sbyte reserve2;
            public byte grp_ident;
            public short user_plen;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0xce)]
            public byte[] user_pdata = new byte[0xce];
            public short cnfg_dlen;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x7e)]
            public byte[] cnfg_data = new byte[0x7e];
            public short slv_ulen;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=30)]
            public byte[] slv_udata = new byte[30];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public byte[] reserve3 = new byte[8];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBSLVPRM3
        {
            public ushort ident_no;
            public byte slv_flag;
            public byte slv_type;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=12)]
            public char[] reserve1 = new char[12];
            public byte slv_stat;
            public byte wd_fact1;
            public byte wd_fact2;
            public byte min_tsdr;
            public char reserve2;
            public byte grp_ident;
            public short user_plen;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0xce)]
            public char[] user_pdata = new char[0xce];
            public short slv_ulen;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=30)]
            public char[] slv_udata = new char[30];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTD
        {
            public short datano;
            public short type;
            public int tool_num;
            public int h_code;
            public int d_code;
            public int tool_inf;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTD2
        {
            public short datano;
            public short dummy;
            public int type;
            public int tool_num;
            public int h_code;
            public int d_code;
            public int tool_inf;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTGI
        {
            public short s_grp;
            public short dummy;
            public short e_grp;
            public FocasLibBase.IODBTGI1 data = new FocasLibBase.IODBTGI1();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTGI_data
        {
            public int n_tool;
            public int count_value;
            public int counter;
            public int count_type;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTGI1
        {
            public FocasLibBase.IODBTGI_data data1 = new FocasLibBase.IODBTGI_data();
            public FocasLibBase.IODBTGI_data data2 = new FocasLibBase.IODBTGI_data();
            public FocasLibBase.IODBTGI_data data3 = new FocasLibBase.IODBTGI_data();
            public FocasLibBase.IODBTGI_data data4 = new FocasLibBase.IODBTGI_data();
            public FocasLibBase.IODBTGI_data data5 = new FocasLibBase.IODBTGI_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTGI2
        {
            public short s_grp;
            public short dummy;
            public short e_grp;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=5)]
            public int[] opt_grpno = new int[5];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTGI3
        {
            public short s_grp;
            public short dummy;
            public short e_grp;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=5)]
            public int[] life_rest = new int[5];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTGI4
        {
            public short grp_no;
            public int n_tool;
            public int count_value;
            public int counter;
            public int count_type;
            public int opt_grpno;
            public int life_rest;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTIME
        {
            public int minute;
            public int msec;
        }

        [StructLayout(LayoutKind.Explicit)]
        public class IODBTIMER
        {
            [FieldOffset(4)]
            public FocasLibBase.TIMER_DATE date;
            [FieldOffset(2)]
            public short dummy;
            [FieldOffset(4)]
            public FocasLibBase.TIMER_TIME time;
            [FieldOffset(0)]
            public short type;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTLCTL
        {
            public short slct;
            public short used_tool;
            public short turret_indx;
            public int zero_tl_no;
            public int t_axis_move;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=2)]
            public int[] total_punch = new int[2];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=11)]
            public short[] reserve = new short[11];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTLDT
        {
            public FocasLibBase.IODBTLDT_data data1 = new FocasLibBase.IODBTLDT_data();
            public FocasLibBase.IODBTLDT_data data2 = new FocasLibBase.IODBTLDT_data();
            public FocasLibBase.IODBTLDT_data data3 = new FocasLibBase.IODBTLDT_data();
            public FocasLibBase.IODBTLDT_data data4 = new FocasLibBase.IODBTLDT_data();
            public FocasLibBase.IODBTLDT_data data5 = new FocasLibBase.IODBTLDT_data();
            public FocasLibBase.IODBTLDT_data data6 = new FocasLibBase.IODBTLDT_data();
            public FocasLibBase.IODBTLDT_data data7 = new FocasLibBase.IODBTLDT_data();
            public FocasLibBase.IODBTLDT_data data8 = new FocasLibBase.IODBTLDT_data();
            public FocasLibBase.IODBTLDT_data data9 = new FocasLibBase.IODBTLDT_data();
            public FocasLibBase.IODBTLDT_data data10 = new FocasLibBase.IODBTLDT_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTLDT_data
        {
            public short slct;
            public int tool_no;
            public int x_axis_ofs;
            public int y_axis_ofs;
            public int turret_pos;
            public int chg_tl_no;
            public int punch_count;
            public int tool_life;
            public int m_tl_radius;
            public int m_tl_angle;
            public byte tl_shape;
            public int tl_size_i;
            public int tl_size_j;
            public int tl_angle;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=3)]
            public int[] reserve = new int[3];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTLGRP
        {
            public FocasLibBase.IODBTLGRP_data data1 = new FocasLibBase.IODBTLGRP_data();
            public FocasLibBase.IODBTLGRP_data data2 = new FocasLibBase.IODBTLGRP_data();
            public FocasLibBase.IODBTLGRP_data data3 = new FocasLibBase.IODBTLGRP_data();
            public FocasLibBase.IODBTLGRP_data data4 = new FocasLibBase.IODBTLGRP_data();
            public FocasLibBase.IODBTLGRP_data data5 = new FocasLibBase.IODBTLGRP_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTLGRP_data
        {
            public int ntool;
            public int nfree;
            public int life;
            public int count;
            public int use_tool;
            public int opt_grpno;
            public int life_rest;
            public short rest_sig;
            public short count_type;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTLMAG
        {
            public FocasLibBase.IODBTLMAG_data data1 = new FocasLibBase.IODBTLMAG_data();
            public FocasLibBase.IODBTLMAG_data data2 = new FocasLibBase.IODBTLMAG_data();
            public FocasLibBase.IODBTLMAG_data data3 = new FocasLibBase.IODBTLMAG_data();
            public FocasLibBase.IODBTLMAG_data data4 = new FocasLibBase.IODBTLMAG_data();
            public FocasLibBase.IODBTLMAG_data data5 = new FocasLibBase.IODBTLMAG_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTLMAG_data
        {
            public short magazine;
            public short pot;
            public short tool_index;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTLMAG2
        {
            public FocasLibBase.IODBTLMAG2_data data1 = new FocasLibBase.IODBTLMAG2_data();
            public FocasLibBase.IODBTLMAG2_data data2 = new FocasLibBase.IODBTLMAG2_data();
            public FocasLibBase.IODBTLMAG2_data data3 = new FocasLibBase.IODBTLMAG2_data();
            public FocasLibBase.IODBTLMAG2_data data4 = new FocasLibBase.IODBTLMAG2_data();
            public FocasLibBase.IODBTLMAG2_data data5 = new FocasLibBase.IODBTLMAG2_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTLMAG2_data
        {
            public short magazine;
            public short pot;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTLMNG
        {
            public FocasLibBase.IODBTLMNG_data data1 = new FocasLibBase.IODBTLMNG_data();
            public FocasLibBase.IODBTLMNG_data data2 = new FocasLibBase.IODBTLMNG_data();
            public FocasLibBase.IODBTLMNG_data data3 = new FocasLibBase.IODBTLMNG_data();
            public FocasLibBase.IODBTLMNG_data data4 = new FocasLibBase.IODBTLMNG_data();
            public FocasLibBase.IODBTLMNG_data data5 = new FocasLibBase.IODBTLMNG_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTLMNG_data
        {
            public int T_code;
            public int life_count;
            public int max_life;
            public int rest_life;
            public byte life_stat;
            public byte cust_bits;
            public ushort tool_info;
            public short H_code;
            public short D_code;
            public int spindle_speed;
            public int feedrate;
            public short magazine;
            public short pot;
            public short gno;
            public short m_ofs;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=4)]
            public int[] reserved = new int[4];
            public int custom1;
            public int custom2;
            public int custom3;
            public int custom4;
            public int custom5;
            public int custom6;
            public int custom7;
            public int custom8;
            public int custom9;
            public int custom10;
            public int custom11;
            public int custom12;
            public int custom13;
            public int custom14;
            public int custom15;
            public int custom16;
            public int custom17;
            public int custom18;
            public int custom19;
            public int custom20;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTLMNG_F2
        {
            public FocasLibBase.IODBTLMNG_F2_data data1 = new FocasLibBase.IODBTLMNG_F2_data();
            public FocasLibBase.IODBTLMNG_F2_data data2 = new FocasLibBase.IODBTLMNG_F2_data();
            public FocasLibBase.IODBTLMNG_F2_data data3 = new FocasLibBase.IODBTLMNG_F2_data();
            public FocasLibBase.IODBTLMNG_F2_data data4 = new FocasLibBase.IODBTLMNG_F2_data();
            public FocasLibBase.IODBTLMNG_F2_data data5 = new FocasLibBase.IODBTLMNG_F2_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTLMNG_F2_data
        {
            public int T_code;
            public int life_count;
            public int max_life;
            public int rest_life;
            public byte life_stat;
            public byte cust_bits;
            public ushort tool_info;
            public short H_code;
            public short D_code;
            public int spindle_speed;
            public int feedrate;
            public short magazine;
            public short pot;
            public short G_code;
            public short W_code;
            public short gno;
            public short m_ofs;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=3)]
            public int[] reserved;
            public int custom1;
            public int custom2;
            public int custom3;
            public int custom4;
            public int custom5;
            public int custom6;
            public int custom7;
            public int custom8;
            public int custom9;
            public int custom10;
            public int custom11;
            public int custom12;
            public int custom13;
            public int custom14;
            public int custom15;
            public int custom16;
            public int custom17;
            public int custom18;
            public int custom19;
            public int custom20;
            public int custom21;
            public int custom22;
            public int custom23;
            public int custom24;
            public int custom25;
            public int custom26;
            public int custom27;
            public int custom28;
            public int custom29;
            public int custom30;
            public int custom31;
            public int custom32;
            public int custom33;
            public int custom34;
            public int custom35;
            public int custom36;
            public int custom37;
            public int custom38;
            public int custom39;
            public int custom40;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTLTOOL
        {
            public FocasLibBase.IODBTLTOOL_data data1 = new FocasLibBase.IODBTLTOOL_data();
            public FocasLibBase.IODBTLTOOL_data data2 = new FocasLibBase.IODBTLTOOL_data();
            public FocasLibBase.IODBTLTOOL_data data3 = new FocasLibBase.IODBTLTOOL_data();
            public FocasLibBase.IODBTLTOOL_data data4 = new FocasLibBase.IODBTLTOOL_data();
            public FocasLibBase.IODBTLTOOL_data data5 = new FocasLibBase.IODBTLTOOL_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTLTOOL_data
        {
            public int tool_num;
            public int h_code;
            public int d_code;
            public int tool_inf;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTLZN
        {
            public FocasLibBase.IODBTLZN_data data1 = new FocasLibBase.IODBTLZN_data();
            public FocasLibBase.IODBTLZN_data data2 = new FocasLibBase.IODBTLZN_data();
            public FocasLibBase.IODBTLZN_data data3 = new FocasLibBase.IODBTLZN_data();
            public FocasLibBase.IODBTLZN_data data4 = new FocasLibBase.IODBTLZN_data();
            public FocasLibBase.IODBTLZN_data data5 = new FocasLibBase.IODBTLZN_data();
            public FocasLibBase.IODBTLZN_data data6 = new FocasLibBase.IODBTLZN_data();
            public FocasLibBase.IODBTLZN_data data7 = new FocasLibBase.IODBTLZN_data();
            public FocasLibBase.IODBTLZN_data data8 = new FocasLibBase.IODBTLZN_data();
            public FocasLibBase.IODBTLZN_data data9 = new FocasLibBase.IODBTLZN_data();
            public FocasLibBase.IODBTLZN_data data10 = new FocasLibBase.IODBTLZN_data();
            public FocasLibBase.IODBTLZN_data data11 = new FocasLibBase.IODBTLZN_data();
            public FocasLibBase.IODBTLZN_data data12 = new FocasLibBase.IODBTLZN_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTLZN_data
        {
            public short slct;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=2)]
            public int[] data = new int[12];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTO_1_1
        {
            public short datano_s;
            public short type;
            public short datano_e;
            public FocasLibBase.OFS_1 ofs = new FocasLibBase.OFS_1();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTO_1_2
        {
            public short datano_s;
            public short type;
            public short datano_e;
            public FocasLibBase.OFS_2 ofs = new FocasLibBase.OFS_2();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTO_1_3
        {
            public short datano_s;
            public short type;
            public short datano_e;
            public FocasLibBase.OFS_3 ofs = new FocasLibBase.OFS_3();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTO_2
        {
            public short datano_s;
            public short type;
            public short datano_e;
            public FocasLibBase.T_OFS_A_data tofsa = new FocasLibBase.T_OFS_A_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTO_3
        {
            public short datano_s;
            public short type;
            public short datano_e;
            public FocasLibBase.T_OFS_B_data tofsb = new FocasLibBase.T_OFS_B_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTR
        {
            public short datano_s;
            public short dummy;
            public short datano_e;
            public FocasLibBase.IODBTR1 data = new FocasLibBase.IODBTR1();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTR_data
        {
            public int ntool;
            public int life;
            public int count;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBTR1
        {
            public FocasLibBase.IODBTR_data data1 = new FocasLibBase.IODBTR_data();
            public FocasLibBase.IODBTR_data data2 = new FocasLibBase.IODBTR_data();
            public FocasLibBase.IODBTR_data data3 = new FocasLibBase.IODBTR_data();
            public FocasLibBase.IODBTR_data data4 = new FocasLibBase.IODBTR_data();
            public FocasLibBase.IODBTR_data data5 = new FocasLibBase.IODBTR_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBUNSOLIC
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x10)]
            public string ipaddr = new string(' ', 0x10);
            public ushort port;
            public short reqaddr;
            public short pmcno;
            public short retry;
            public short timeout;
            public short alivetime;
            public short setno;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=3)]
            public FocasLibBase.IODBUNSOLIC_pmc[] rddata = new FocasLibBase.IODBUNSOLIC_pmc[3];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public struct IODBUNSOLIC_pmc
        {
            public short type;
            public short rdaddr;
            public short rdno;
            public short rdsize;
            private int dummy;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBWAVE
        {
            public short condition;
            public char trg_adr;
            public byte trg_bit;
            public short trg_no;
            public short delay;
            public short t_range;
            public FocasLibBase.IODBWAVE_ch ch = new FocasLibBase.IODBWAVE_ch();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBWAVE_axis
        {
            public short axis;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBWAVE_ch
        {
            public FocasLibBase.IODBWAVE_ch_data ch1 = new FocasLibBase.IODBWAVE_ch_data();
            public FocasLibBase.IODBWAVE_ch_data ch2 = new FocasLibBase.IODBWAVE_ch_data();
            public FocasLibBase.IODBWAVE_ch_data ch3 = new FocasLibBase.IODBWAVE_ch_data();
            public FocasLibBase.IODBWAVE_ch_data ch4 = new FocasLibBase.IODBWAVE_ch_data();
            public FocasLibBase.IODBWAVE_ch_data ch5 = new FocasLibBase.IODBWAVE_ch_data();
            public FocasLibBase.IODBWAVE_ch_data ch6 = new FocasLibBase.IODBWAVE_ch_data();
            public FocasLibBase.IODBWAVE_ch_data ch7 = new FocasLibBase.IODBWAVE_ch_data();
            public FocasLibBase.IODBWAVE_ch_data ch8 = new FocasLibBase.IODBWAVE_ch_data();
            public FocasLibBase.IODBWAVE_ch_data ch9 = new FocasLibBase.IODBWAVE_ch_data();
            public FocasLibBase.IODBWAVE_ch_data ch10 = new FocasLibBase.IODBWAVE_ch_data();
            public FocasLibBase.IODBWAVE_ch_data ch11 = new FocasLibBase.IODBWAVE_ch_data();
            public FocasLibBase.IODBWAVE_ch_data ch12 = new FocasLibBase.IODBWAVE_ch_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBWAVE_ch_data
        {
            public short kind;
            public FocasLibBase.IODBWAVE_u u = new FocasLibBase.IODBWAVE_u();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBWAVE_io
        {
            public byte adr;
            public byte bit;
            public short no;
        }

        [StructLayout(LayoutKind.Explicit)]
        public class IODBWAVE_u
        {
            [FieldOffset(0)]
            public FocasLibBase.IODBWAVE_axis axis = new FocasLibBase.IODBWAVE_axis();
            [FieldOffset(0)]
            public FocasLibBase.IODBWAVE_io io = new FocasLibBase.IODBWAVE_io();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBWCSF
        {
            public short datano;
            public short type;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] data = new int[8];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBWVPRM
        {
            public short condition;
            public byte trg_adr;
            public byte trg_bit;
            public short trg_no;
            public short reserve1;
            public int delay;
            public int t_range;
            public FocasLibBase.IODBWVPRM_ch ch = new FocasLibBase.IODBWVPRM_ch();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBWVPRM_axis
        {
            public short axis;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBWVPRM_ch
        {
            public FocasLibBase.IODBWVPRM_ch_data ch1 = new FocasLibBase.IODBWVPRM_ch_data();
            public FocasLibBase.IODBWVPRM_ch_data ch2 = new FocasLibBase.IODBWVPRM_ch_data();
            public FocasLibBase.IODBWVPRM_ch_data ch3 = new FocasLibBase.IODBWVPRM_ch_data();
            public FocasLibBase.IODBWVPRM_ch_data ch4 = new FocasLibBase.IODBWVPRM_ch_data();
            public FocasLibBase.IODBWVPRM_ch_data ch5 = new FocasLibBase.IODBWVPRM_ch_data();
            public FocasLibBase.IODBWVPRM_ch_data ch6 = new FocasLibBase.IODBWVPRM_ch_data();
            public FocasLibBase.IODBWVPRM_ch_data ch7 = new FocasLibBase.IODBWVPRM_ch_data();
            public FocasLibBase.IODBWVPRM_ch_data ch8 = new FocasLibBase.IODBWVPRM_ch_data();
            public FocasLibBase.IODBWVPRM_ch_data ch9 = new FocasLibBase.IODBWVPRM_ch_data();
            public FocasLibBase.IODBWVPRM_ch_data ch10 = new FocasLibBase.IODBWVPRM_ch_data();
            public FocasLibBase.IODBWVPRM_ch_data ch11 = new FocasLibBase.IODBWVPRM_ch_data();
            public FocasLibBase.IODBWVPRM_ch_data ch12 = new FocasLibBase.IODBWVPRM_ch_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBWVPRM_ch_data
        {
            public short kind;
            public FocasLibBase.IODBWVPRM_u u = new FocasLibBase.IODBWVPRM_u();
            public int reserve2;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBWVPRM_io
        {
            public byte adr;
            public byte bit;
            public short no;
        }

        [StructLayout(LayoutKind.Explicit)]
        public class IODBWVPRM_u
        {
            [FieldOffset(0)]
            public FocasLibBase.IODBWVPRM_axis axis = new FocasLibBase.IODBWVPRM_axis();
            [FieldOffset(0)]
            public FocasLibBase.IODBWVPRM_io io = new FocasLibBase.IODBWVPRM_io();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBZOFS
        {
            public short datano;
            public short type;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] data = new int[8];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class IODBZOR
        {
            public short datano_s;
            public short type;
            public short datano_e;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x38)]
            public int[] data = new int[0x38];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class LOADELM
        {
            public int data;
            public short dec;
            public short unit;
            public byte name;
            public byte suff1;
            public byte suff2;
            public byte reserve;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class M_CODE_data
        {
            public int no;
            public short flag;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class M_CODE1
        {
            public FocasLibBase.M_CODE_data m_code1 = new FocasLibBase.M_CODE_data();
            public FocasLibBase.M_CODE_data m_code2 = new FocasLibBase.M_CODE_data();
            public FocasLibBase.M_CODE_data m_code3 = new FocasLibBase.M_CODE_data();
            public FocasLibBase.M_CODE_data m_code4 = new FocasLibBase.M_CODE_data();
            public FocasLibBase.M_CODE_data m_code5 = new FocasLibBase.M_CODE_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class MODAL_AUX_data
        {
            public int aux_data;
            public byte flag1;
            public byte flag2;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class MODAL_RAUX1_data
        {
            public FocasLibBase.MODAL_AUX_data data1 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data2 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data3 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data4 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data5 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data6 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data7 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data8 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data9 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data10 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data11 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data12 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data13 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data14 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data15 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data16 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data17 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data18 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data19 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data20 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data21 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data22 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data23 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data24 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data25 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data26 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data27 = new FocasLibBase.MODAL_AUX_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class MODAL_RAUX2_data
        {
            public FocasLibBase.MODAL_AUX_data data1 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data2 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data3 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data4 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data5 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data6 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data7 = new FocasLibBase.MODAL_AUX_data();
            public FocasLibBase.MODAL_AUX_data data8 = new FocasLibBase.MODAL_AUX_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class OAXIS
        {
            public int absolute;
            public int machine;
            public int relative;
            public int distance;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODB3DCD
        {
            public short mode;
            public short dno;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=3)]
            public short[] cd_axes;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=6)]
            public int[] center;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=6)]
            public int[] direct;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=2)]
            public int[] angle;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODB3DHDL
        {
            public FocasLibBase.ODB3DHDL_data data1 = new FocasLibBase.ODB3DHDL_data();
            public FocasLibBase.ODB3DHDL_data data2 = new FocasLibBase.ODB3DHDL_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODB3DHDL_data
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=5)]
            public short[] axes;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=5)]
            public int[] data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODB3DPLS
        {
            public FocasLibBase.ODB3DPLS_data pls1 = new FocasLibBase.ODB3DPLS_data();
            public FocasLibBase.ODB3DPLS_data pls2 = new FocasLibBase.ODB3DPLS_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODB3DPLS_data
        {
            public int right_angle_x;
            public int right_angle_y;
            public int tool_axis;
            public int tool_tip_a_b;
            public int tool_tip_c;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODB3DTO
        {
            public short mode;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=3)]
            public short[] ofs_axes;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=3)]
            public int[] ofsvct;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODB5AXMAN
        {
            public short type1;
            public short type2;
            public short type3;
            public int data1;
            public int data2;
            public int data3;
            public int c1;
            public int c2;
            public int dummy;
            public int td;
            public int r1;
            public int r2;
            public int vr;
            public int h1;
            public int h2;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBACT
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=2)]
            public short[] dummy;
            public int data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBACT2
        {
            public short datano;
            public short type;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=4)]
            public int[] data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBACTTLZN
        {
            public short act_no;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=2)]
            public int[] data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBAD
        {
            public short datano;
            public short type;
            public short data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBAHIS
        {
            public ushort s_no;
            public short type;
            public ushort e_no;
            public FocasLibBase.ALM_HIS1 alm_his = new FocasLibBase.ALM_HIS1();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBAHIS2
        {
            public ushort s_no;
            public ushort e_no;
            public FocasLibBase.ALM_HIS2 alm_his = new FocasLibBase.ALM_HIS2();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBAHIS3
        {
            public ushort s_no;
            public ushort e_no;
            public FocasLibBase.ALM_HIS3 alm_his = new FocasLibBase.ALM_HIS3();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBAHIS5
        {
            public ushort s_no;
            public ushort e_no;
            public FocasLibBase.ALM_HIS5 alm_his = new FocasLibBase.ALM_HIS5();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBALM
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=2)]
            public short[] dummy = new short[2];
            public ushort data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBALMMSG
        {
            public FocasLibBase.ODBALMMSG_data msg1 = new FocasLibBase.ODBALMMSG_data();
            public FocasLibBase.ODBALMMSG_data msg2 = new FocasLibBase.ODBALMMSG_data();
            public FocasLibBase.ODBALMMSG_data msg3 = new FocasLibBase.ODBALMMSG_data();
            public FocasLibBase.ODBALMMSG_data msg4 = new FocasLibBase.ODBALMMSG_data();
            public FocasLibBase.ODBALMMSG_data msg5 = new FocasLibBase.ODBALMMSG_data();
            public FocasLibBase.ODBALMMSG_data msg6 = new FocasLibBase.ODBALMMSG_data();
            public FocasLibBase.ODBALMMSG_data msg7 = new FocasLibBase.ODBALMMSG_data();
            public FocasLibBase.ODBALMMSG_data msg8 = new FocasLibBase.ODBALMMSG_data();
            public FocasLibBase.ODBALMMSG_data msg9 = new FocasLibBase.ODBALMMSG_data();
            public FocasLibBase.ODBALMMSG_data msg10 = new FocasLibBase.ODBALMMSG_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBALMMSG_data
        {
            public int alm_no;
            public short type;
            public short axis;
            public short dummy;
            public short msg_len;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x20)]
            public string alm_msg = new string(' ', 0x20);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBALMMSG2
        {
            public FocasLibBase.ODBALMMSG2_data msg1 = new FocasLibBase.ODBALMMSG2_data();
            public FocasLibBase.ODBALMMSG2_data msg2 = new FocasLibBase.ODBALMMSG2_data();
            public FocasLibBase.ODBALMMSG2_data msg3 = new FocasLibBase.ODBALMMSG2_data();
            public FocasLibBase.ODBALMMSG2_data msg4 = new FocasLibBase.ODBALMMSG2_data();
            public FocasLibBase.ODBALMMSG2_data msg5 = new FocasLibBase.ODBALMMSG2_data();
            public FocasLibBase.ODBALMMSG2_data msg6 = new FocasLibBase.ODBALMMSG2_data();
            public FocasLibBase.ODBALMMSG2_data msg7 = new FocasLibBase.ODBALMMSG2_data();
            public FocasLibBase.ODBALMMSG2_data msg8 = new FocasLibBase.ODBALMMSG2_data();
            public FocasLibBase.ODBALMMSG2_data msg9 = new FocasLibBase.ODBALMMSG2_data();
            public FocasLibBase.ODBALMMSG2_data msg10 = new FocasLibBase.ODBALMMSG2_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBALMMSG2_data
        {
            public int alm_no;
            public short type;
            public short axis;
            public short dummy;
            public short msg_len;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x40)]
            public string alm_msg = new string(' ', 0x40);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBAXDT
        {
            public FocasLibBase.ODBAXDT_data data1 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data2 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data3 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data4 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data5 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data6 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data7 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data8 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data9 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data10 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data11 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data12 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data13 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data14 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data15 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data16 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data17 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data18 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data19 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data20 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data21 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data22 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data23 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data24 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data25 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data26 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data27 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data28 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data29 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data30 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data31 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data32 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data33 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data34 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data35 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data36 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data37 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data38 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data39 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data40 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data41 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data42 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data43 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data44 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data45 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data46 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data47 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data48 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data49 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data50 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data51 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data52 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data53 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data54 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data55 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data56 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data57 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data58 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data59 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data60 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data61 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data62 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data63 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data64 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data65 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data66 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data67 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data68 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data69 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data70 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data71 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data72 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data73 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data74 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data75 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data76 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data77 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data78 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data79 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data80 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data81 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data82 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data83 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data84 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data85 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data86 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data87 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data88 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data89 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data90 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data91 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data92 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data93 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data94 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data95 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data96 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data97 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data98 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data99 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data100 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data101 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data102 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data103 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data104 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data105 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data106 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data107 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data108 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data109 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data110 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data111 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data112 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data113 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data114 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data115 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data116 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data117 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data118 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data119 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data120 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data121 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data122 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data123 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data124 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data125 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data126 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data127 = new FocasLibBase.ODBAXDT_data();
            public FocasLibBase.ODBAXDT_data data128 = new FocasLibBase.ODBAXDT_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBAXDT_data
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string name = new string(' ', 4);
            public int data;
            public short dec;
            public short unit;
            public short flag;
            public short reserve;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBAXIS
        {
            public short dummy;
            public short type;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBAXISNAME
        {
            public FocasLibBase.ODBAXISNAME_data data1 = new FocasLibBase.ODBAXISNAME_data();
            public FocasLibBase.ODBAXISNAME_data data2 = new FocasLibBase.ODBAXISNAME_data();
            public FocasLibBase.ODBAXISNAME_data data3 = new FocasLibBase.ODBAXISNAME_data();
            public FocasLibBase.ODBAXISNAME_data data4 = new FocasLibBase.ODBAXISNAME_data();
            public FocasLibBase.ODBAXISNAME_data data5 = new FocasLibBase.ODBAXISNAME_data();
            public FocasLibBase.ODBAXISNAME_data data6 = new FocasLibBase.ODBAXISNAME_data();
            public FocasLibBase.ODBAXISNAME_data data7 = new FocasLibBase.ODBAXISNAME_data();
            public FocasLibBase.ODBAXISNAME_data data8 = new FocasLibBase.ODBAXISNAME_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBAXISNAME_data
        {
            public byte name;
            public byte suff;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBBAXIS
        {
            public short flag;
            public short command;
            public ushort speed;
            public int sub_data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBBRS
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] dest;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] dist;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBBTLINF
        {
            public short ofs_type;
            public short use_no;
            public short sub_no;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBBUF
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=2)]
            public short[] dummy;
            public short data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBCAXIS
        {
            public short dummy;
            public short type;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public sbyte[] data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBCMD
        {
            public FocasLibBase.ODBCMD_data cmd0 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd1 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd2 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd3 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd4 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd5 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd6 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd7 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd8 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd9 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd10 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd11 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd12 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd13 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd14 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd15 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd16 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd17 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd18 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd19 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd20 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd21 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd22 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd23 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd24 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd25 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd26 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd27 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd28 = new FocasLibBase.ODBCMD_data();
            public FocasLibBase.ODBCMD_data cmd29 = new FocasLibBase.ODBCMD_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBCMD_data
        {
            public byte adrs;
            public byte num;
            public short flag;
            public int cmd_val;
            public int dec_val;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBCSS
        {
            public int srpm;
            public int sspm;
            public int smax;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBDATRNG
        {
            public int data_min;
            public int data_max;
            public int status;
        }

        [StructLayout(LayoutKind.Explicit)]
        public class ODBDGN_1
        {
            [FieldOffset(4)]
            public byte cdata;
            [FieldOffset(0)]
            public short datano;
            [FieldOffset(4)]
            public short idata;
            [FieldOffset(4)]
            public int ldata;
            [FieldOffset(2)]
            public short type;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBDGN_2
        {
            public short datano;
            public short type;
            public FocasLibBase.REALDGN rdata = new FocasLibBase.REALDGN();
        }

        [StructLayout(LayoutKind.Explicit)]
        public class ODBDGN_3
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8), FieldOffset(4)]
            public byte[] cdatas;
            [FieldOffset(0)]
            public short datano;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8), FieldOffset(4)]
            public short[] idatas;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8), FieldOffset(4)]
            public int[] ldatas;
            [FieldOffset(2)]
            public short type;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBDGN_4
        {
            public short datano;
            public short type;
            public FocasLibBase.REALDGNS rdatas = new FocasLibBase.REALDGNS();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBDGN_A
        {
            public FocasLibBase.ODBDGN_1 data1 = new FocasLibBase.ODBDGN_1();
            public FocasLibBase.ODBDGN_1 data2 = new FocasLibBase.ODBDGN_1();
            public FocasLibBase.ODBDGN_1 data3 = new FocasLibBase.ODBDGN_1();
            public FocasLibBase.ODBDGN_1 data4 = new FocasLibBase.ODBDGN_1();
            public FocasLibBase.ODBDGN_1 data5 = new FocasLibBase.ODBDGN_1();
            public FocasLibBase.ODBDGN_1 data6 = new FocasLibBase.ODBDGN_1();
            public FocasLibBase.ODBDGN_1 data7 = new FocasLibBase.ODBDGN_1();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBDGN_B
        {
            public FocasLibBase.ODBDGN_2 data1 = new FocasLibBase.ODBDGN_2();
            public FocasLibBase.ODBDGN_2 data2 = new FocasLibBase.ODBDGN_2();
            public FocasLibBase.ODBDGN_2 data3 = new FocasLibBase.ODBDGN_2();
            public FocasLibBase.ODBDGN_2 data4 = new FocasLibBase.ODBDGN_2();
            public FocasLibBase.ODBDGN_2 data5 = new FocasLibBase.ODBDGN_2();
            public FocasLibBase.ODBDGN_2 data6 = new FocasLibBase.ODBDGN_2();
            public FocasLibBase.ODBDGN_2 data7 = new FocasLibBase.ODBDGN_2();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBDGN_C
        {
            public FocasLibBase.ODBDGN_3 data1 = new FocasLibBase.ODBDGN_3();
            public FocasLibBase.ODBDGN_3 data2 = new FocasLibBase.ODBDGN_3();
            public FocasLibBase.ODBDGN_3 data3 = new FocasLibBase.ODBDGN_3();
            public FocasLibBase.ODBDGN_3 data4 = new FocasLibBase.ODBDGN_3();
            public FocasLibBase.ODBDGN_3 data5 = new FocasLibBase.ODBDGN_3();
            public FocasLibBase.ODBDGN_3 data6 = new FocasLibBase.ODBDGN_3();
            public FocasLibBase.ODBDGN_3 data7 = new FocasLibBase.ODBDGN_3();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBDGN_D
        {
            public FocasLibBase.ODBDGN_4 data1 = new FocasLibBase.ODBDGN_4();
            public FocasLibBase.ODBDGN_4 data2 = new FocasLibBase.ODBDGN_4();
            public FocasLibBase.ODBDGN_4 data3 = new FocasLibBase.ODBDGN_4();
            public FocasLibBase.ODBDGN_4 data4 = new FocasLibBase.ODBDGN_4();
            public FocasLibBase.ODBDGN_4 data5 = new FocasLibBase.ODBDGN_4();
            public FocasLibBase.ODBDGN_4 data6 = new FocasLibBase.ODBDGN_4();
            public FocasLibBase.ODBDGN_4 data7 = new FocasLibBase.ODBDGN_4();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBDIAGIF
        {
            public ushort info_no;
            public short prev_no;
            public short next_no;
            public FocasLibBase.ODBDIAGIF1 info = new FocasLibBase.ODBDIAGIF1();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBDIAGIF_info
        {
            public short diag_no;
            public short diag_type;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBDIAGIF1
        {
            public FocasLibBase.ODBDIAGIF_info info1 = new FocasLibBase.ODBDIAGIF_info();
            public FocasLibBase.ODBDIAGIF_info info2 = new FocasLibBase.ODBDIAGIF_info();
            public FocasLibBase.ODBDIAGIF_info info3 = new FocasLibBase.ODBDIAGIF_info();
            public FocasLibBase.ODBDIAGIF_info info4 = new FocasLibBase.ODBDIAGIF_info();
            public FocasLibBase.ODBDIAGIF_info info5 = new FocasLibBase.ODBDIAGIF_info();
            public FocasLibBase.ODBDIAGIF_info info6 = new FocasLibBase.ODBDIAGIF_info();
            public FocasLibBase.ODBDIAGIF_info info7 = new FocasLibBase.ODBDIAGIF_info();
            public FocasLibBase.ODBDIAGIF_info info8 = new FocasLibBase.ODBDIAGIF_info();
            public FocasLibBase.ODBDIAGIF_info info9 = new FocasLibBase.ODBDIAGIF_info();
            public FocasLibBase.ODBDIAGIF_info info10 = new FocasLibBase.ODBDIAGIF_info();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBDIAGNUM
        {
            public ushort diag_min;
            public ushort diag_max;
            public ushort total_no;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBDISCHRG
        {
            public ushort aps;
            public ushort year;
            public ushort month;
            public ushort day;
            public ushort hour;
            public ushort minute;
            public ushort second;
            public short hpc;
            public short hfq;
            public short hdt;
            public short hpa;
            public int hce;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] rfi;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] rfv;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] dci;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] dcv;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] dcw;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBDISCHRGALM
        {
            public FocasLibBase.ODBDISCHRGALM_data data1 = new FocasLibBase.ODBDISCHRGALM_data();
            public FocasLibBase.ODBDISCHRGALM_data data2 = new FocasLibBase.ODBDISCHRGALM_data();
            public FocasLibBase.ODBDISCHRGALM_data data3 = new FocasLibBase.ODBDISCHRGALM_data();
            public FocasLibBase.ODBDISCHRGALM_data data4 = new FocasLibBase.ODBDISCHRGALM_data();
            public FocasLibBase.ODBDISCHRGALM_data data5 = new FocasLibBase.ODBDISCHRGALM_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBDISCHRGALM_data
        {
            public ushort year;
            public ushort month;
            public ushort day;
            public ushort hour;
            public ushort minute;
            public ushort second;
            public int almnum;
            public uint psec;
            public short hpc;
            public short hfq;
            public short hdt;
            public short hpa;
            public int hce;
            public ushort asq;
            public ushort psu;
            public ushort aps;
            public short dummy;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] rfi;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] rfv;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] dci;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] dcv;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] dcw;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public short[] almcd;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBDNCDGN
        {
            public short ctrl_word;
            public short can_word;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x10)]
            public char[] nc_file;
            public ushort read_ptr;
            public ushort write_ptr;
            public ushort empty_cnt;
            public uint total_size;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBDSDIR
        {
            public int file_num;
            public int remainder;
            public short data_num;
            public FocasLibBase.ODBDSDIR1 data = new FocasLibBase.ODBDSDIR1();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBDSDIR_data
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x10)]
            public string file_name = new string(' ', 0x10);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x40)]
            public string comment = new string(' ', 0x40);
            public int size;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x10)]
            public string date = new string(' ', 0x10);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBDSDIR1
        {
            public FocasLibBase.ODBDSDIR_data data1 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data2 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data3 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data4 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data5 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data6 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data7 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data8 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data9 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data10 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data11 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data12 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data13 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data14 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data15 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data16 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data17 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data18 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data19 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data20 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data21 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data22 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data23 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data24 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data25 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data26 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data27 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data28 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data29 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data30 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data31 = new FocasLibBase.ODBDSDIR_data();
            public FocasLibBase.ODBDSDIR_data data32 = new FocasLibBase.ODBDSDIR_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBDSMNT
        {
            public int empty_cnt;
            public int total_size;
            public int read_ptr;
            public int write_ptr;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBDY_1
        {
            public short dummy;
            public short axis;
            public short alarm;
            public short prgnum;
            public short prgmnum;
            public int seqnum;
            public int actf;
            public int acts;
            public FocasLibBase.FAXIS pos = new FocasLibBase.FAXIS();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBDY_2
        {
            public short dummy;
            public short axis;
            public short alarm;
            public short prgnum;
            public short prgmnum;
            public int seqnum;
            public int actf;
            public int acts;
            public FocasLibBase.OAXIS pos = new FocasLibBase.OAXIS();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBDY2_1
        {
            public short dummy;
            public short axis;
            public int alarm;
            public int prgnum;
            public int prgmnum;
            public int seqnum;
            public int actf;
            public int acts;
            public FocasLibBase.FAXIS pos = new FocasLibBase.FAXIS();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBDY2_2
        {
            public short dummy;
            public short axis;
            public int alarm;
            public int prgnum;
            public int prgmnum;
            public int seqnum;
            public int actf;
            public int acts;
            public FocasLibBase.OAXIS pos = new FocasLibBase.OAXIS();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBERR
        {
            public short err_no;
            public short err_dtno;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBETMSG
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x21)]
            public string title = new string(' ', 0x21);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=390)]
            public string message = new string(' ', 390);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBEXAXISNAME
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname1 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname2 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname3 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname4 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname5 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname6 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname7 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname8 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname9 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname10 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname11 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname12 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname13 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname14 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname15 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname16 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname17 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname18 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname19 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname20 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname21 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname22 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname23 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname24 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname25 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname26 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname27 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname28 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname29 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname30 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname31 = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string axname32 = new string(' ', 4);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBEXEM
        {
            public short grp_no;
            public short mem_no;
            public FocasLibBase.ODBEXEM1 m_code = new FocasLibBase.ODBEXEM1();
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x15)]
            public string m_name = new string(' ', 0x15);
            public byte dummy;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBEXEM_data
        {
            public int no;
            public short flag;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBEXEM1
        {
            public FocasLibBase.ODBEXEM_data m_code1 = new FocasLibBase.ODBEXEM_data();
            public FocasLibBase.ODBEXEM_data m_code2 = new FocasLibBase.ODBEXEM_data();
            public FocasLibBase.ODBEXEM_data m_code3 = new FocasLibBase.ODBEXEM_data();
            public FocasLibBase.ODBEXEM_data m_code4 = new FocasLibBase.ODBEXEM_data();
            public FocasLibBase.ODBEXEM_data m_code5 = new FocasLibBase.ODBEXEM_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBEXEPRG
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x24)]
            public char[] name;
            public int o_num;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBEXGP
        {
            public FocasLibBase.ODBEXGP_data data1 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data2 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data3 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data4 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data5 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data6 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data7 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data8 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data9 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data10 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data11 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data12 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data13 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data14 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data15 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data16 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data17 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data18 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data19 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data20 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data21 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data22 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data23 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data24 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data25 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data26 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data27 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data28 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data29 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data30 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data31 = new FocasLibBase.ODBEXGP_data();
            public FocasLibBase.ODBEXGP_data data32 = new FocasLibBase.ODBEXGP_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBEXGP_data
        {
            public int grp_no;
            public int opt_grpno;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBFINFO
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=12)]
            public string slotname = new string(' ', 12);
            public int fromnum;
            public FocasLibBase.ODBFINFO1 info = new FocasLibBase.ODBFINFO1();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBFINFO_info
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=12)]
            public string sysname = new string(' ', 12);
            public int fromsize;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBFINFO1
        {
            public FocasLibBase.ODBFINFO_info info1 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info2 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info3 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info4 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info5 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info6 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info7 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info8 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info9 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info10 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info11 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info12 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info13 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info14 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info15 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info16 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info17 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info18 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info19 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info20 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info21 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info22 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info23 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info24 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info25 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info26 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info27 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info28 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info29 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info30 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info31 = new FocasLibBase.ODBFINFO_info();
            public FocasLibBase.ODBFINFO_info info32 = new FocasLibBase.ODBFINFO_info();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBFINFORM
        {
            public int slotno;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=12)]
            public string slotname = new string(' ', 12);
            public int fromnum;
            public FocasLibBase.ODBFINFORM1 info = new FocasLibBase.ODBFINFORM1();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBFINFORM_info
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=12)]
            public string sysname = new string(' ', 12);
            public int fromsize;
            public int fromattrib;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBFINFORM1
        {
            public FocasLibBase.ODBFINFORM_info info1 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info2 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info3 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info4 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info5 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info6 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info7 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info8 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info9 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info10 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info11 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info12 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info13 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info14 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info15 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info16 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info17 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info18 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info19 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info20 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info21 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info22 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info23 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info24 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info25 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info26 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info27 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info28 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info29 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info30 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info31 = new FocasLibBase.ODBFINFORM_info();
            public FocasLibBase.ODBFINFORM_info info32 = new FocasLibBase.ODBFINFORM_info();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBFIX
        {
            public short mode;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=2)]
            public short[] pln_axes;
            public short drl_axes;
            public int i_pos;
            public int r_pos;
            public int z_pos;
            public int cmd_cnt;
            public int act_cnt;
            public int cut;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=2)]
            public int[] shift;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBGCD
        {
            public FocasLibBase.ODBGCD_data gcd0 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd1 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd2 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd3 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd4 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd5 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd6 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd7 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd8 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd9 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd10 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd11 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd12 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd13 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd14 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd15 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd16 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd17 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd18 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd19 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd20 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd21 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd22 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd23 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd24 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd25 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd26 = new FocasLibBase.ODBGCD_data();
            public FocasLibBase.ODBGCD_data gcd27 = new FocasLibBase.ODBGCD_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBGCD_data
        {
            public short group;
            public short flag;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=8)]
            public string code = new string(' ', 8);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBHDDDIR
        {
            public FocasLibBase.ODBHDDDIR_data data1 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data2 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data3 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data4 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data5 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data6 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data7 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data8 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data9 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data10 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data11 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data12 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data13 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data14 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data15 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data16 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data17 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data18 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data19 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data20 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data21 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data22 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data23 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data24 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data25 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data26 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data27 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data28 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data29 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data30 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data31 = new FocasLibBase.ODBHDDDIR_data();
            public FocasLibBase.ODBHDDDIR_data data32 = new FocasLibBase.ODBHDDDIR_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBHDDDIR_data
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x40)]
            public string file_name = new string(' ', 0x40);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=80)]
            public string comment = new string(' ', 80);
            public short attribute;
            public short reserved;
            public int size;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x10)]
            public string date = new string(' ', 0x10);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBHDDINF
        {
            public int file_num;
            public int remainder_l;
            public int remainder_h;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x20)]
            public char[] current_dir = new char[0x20];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBHIS
        {
            public ushort s_no;
            public short type;
            public ushort e_no;
            public FocasLibBase.ODBHIS1 data = new FocasLibBase.ODBHIS1();
        }

        [StructLayout(LayoutKind.Explicit, Size=8)]
        public class ODBHIS_data
        {
            [FieldOffset(2)]
            public short alm_alm_grp;
            [FieldOffset(4)]
            public short alm_alm_no;
            [FieldOffset(6)]
            public sbyte alm_axis_no;
            [FieldOffset(7)]
            public sbyte alm_dummy;
            [FieldOffset(0)]
            public short alm_rec_type;
            [FieldOffset(4)]
            public sbyte date_day;
            [FieldOffset(6)]
            public sbyte date_dummy1;
            [FieldOffset(7)]
            public sbyte date_dummy2;
            [FieldOffset(3)]
            public sbyte date_month;
            [FieldOffset(5)]
            public sbyte date_pw_flag;
            [FieldOffset(0)]
            public short date_rec_type;
            [FieldOffset(2)]
            public sbyte date_year;
            [FieldOffset(4)]
            public sbyte mdi_dummy1;
            [FieldOffset(5)]
            public sbyte mdi_dummy2;
            [FieldOffset(6)]
            public sbyte mdi_dummy3;
            [FieldOffset(7)]
            public sbyte mdi_dummy4;
            [FieldOffset(2)]
            public byte mdi_key_code;
            [FieldOffset(3)]
            public byte mdi_pw_flag;
            [FieldOffset(0)]
            public short mdi_rec_type;
            [FieldOffset(0)]
            public short rec_type;
            [FieldOffset(5)]
            public sbyte sgn_dummy;
            [FieldOffset(0)]
            public short sgn_rec_type;
            [FieldOffset(2)]
            public sbyte sgn_sig_name;
            [FieldOffset(4)]
            public byte sgn_sig_new;
            [FieldOffset(6)]
            public short sgn_sig_no;
            [FieldOffset(3)]
            public byte sgn_sig_old;
            [FieldOffset(6)]
            public sbyte time_dummy1;
            [FieldOffset(7)]
            public sbyte time_dummy2;
            [FieldOffset(2)]
            public sbyte time_hour;
            [FieldOffset(3)]
            public sbyte time_minute;
            [FieldOffset(5)]
            public sbyte time_pw_flag;
            [FieldOffset(0)]
            public short time_rec_type;
            [FieldOffset(4)]
            public sbyte time_second;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBHIS1
        {
            public FocasLibBase.ODBHIS_data data1 = new FocasLibBase.ODBHIS_data();
            public FocasLibBase.ODBHIS_data data2 = new FocasLibBase.ODBHIS_data();
            public FocasLibBase.ODBHIS_data data3 = new FocasLibBase.ODBHIS_data();
            public FocasLibBase.ODBHIS_data data4 = new FocasLibBase.ODBHIS_data();
            public FocasLibBase.ODBHIS_data data5 = new FocasLibBase.ODBHIS_data();
            public FocasLibBase.ODBHIS_data data6 = new FocasLibBase.ODBHIS_data();
            public FocasLibBase.ODBHIS_data data7 = new FocasLibBase.ODBHIS_data();
            public FocasLibBase.ODBHIS_data data8 = new FocasLibBase.ODBHIS_data();
            public FocasLibBase.ODBHIS_data data9 = new FocasLibBase.ODBHIS_data();
            public FocasLibBase.ODBHIS_data data10 = new FocasLibBase.ODBHIS_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBHND
        {
            public FocasLibBase.ODBHND_data p1 = new FocasLibBase.ODBHND_data();
            public FocasLibBase.ODBHND_data p2 = new FocasLibBase.ODBHND_data();
            public FocasLibBase.ODBHND_data p3 = new FocasLibBase.ODBHND_data();
            public FocasLibBase.ODBHND_data p4 = new FocasLibBase.ODBHND_data();
            public FocasLibBase.ODBHND_data p5 = new FocasLibBase.ODBHND_data();
            public FocasLibBase.ODBHND_data p6 = new FocasLibBase.ODBHND_data();
            public FocasLibBase.ODBHND_data p7 = new FocasLibBase.ODBHND_data();
            public FocasLibBase.ODBHND_data p8 = new FocasLibBase.ODBHND_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBHND_data
        {
            public FocasLibBase.POSELM input = new FocasLibBase.POSELM();
            public FocasLibBase.POSELM output = new FocasLibBase.POSELM();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBHOSTDIR
        {
            public FocasLibBase.ODBHOSTDIR_data data1 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data2 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data3 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data4 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data5 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data6 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data7 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data8 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data9 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data10 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data11 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data12 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data13 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data14 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data15 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data16 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data17 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data18 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data19 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data20 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data21 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data22 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data23 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data24 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data25 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data26 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data27 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data28 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data29 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data30 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data31 = new FocasLibBase.ODBHOSTDIR_data();
            public FocasLibBase.ODBHOSTDIR_data data32 = new FocasLibBase.ODBHOSTDIR_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBHOSTDIR_data
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x80)]
            public char[] host_file = new char[0x80];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBLACTN
        {
            public short slct;
            public short act_proc;
            public short act_pirce;
            public short act_slop;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=5)]
            public short[] reserve = new short[5];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBLCMDT
        {
            public short slct;
            public int feed;
            public short power;
            public short freq;
            public short duty;
            public short g_kind;
            public short g_ready_t;
            public short g_press;
            public short error;
            public int dsplc;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=7)]
            public short[] reserve = new short[7];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBLCMMT
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x19)]
            public string comment = new string(' ', 0x19);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBLFNO
        {
            public short datano;
            public short type;
            public short data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBLOFS
        {
            public short mode;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] ofsvct;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBLOPDT
        {
            public short slct;
            public short pwr_mon;
            public short pwr_ofs;
            public short pwr_act;
            public int feed_act;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=4)]
            public short[] reserve;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBM
        {
            public short datano;
            public short dummy;
            public int mcr_val;
            public short dec_val;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBMDIP
        {
            public short mdiprog;
            public int mdipntr;
            public short crntprog;
            public int crntpntr;
        }

        [StructLayout(LayoutKind.Explicit)]
        public class ODBMDL_1
        {
            [FieldOffset(0)]
            public short datano;
            [FieldOffset(4)]
            public byte g_data;
            [FieldOffset(2)]
            public short type;
        }

        [StructLayout(LayoutKind.Explicit)]
        public class ODBMDL_2
        {
            [FieldOffset(0)]
            public short datano;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=4), FieldOffset(4)]
            public byte[] g_1shot = new byte[4];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x23), FieldOffset(4)]
            public byte[] g_rdata = new byte[0x23];
            [FieldOffset(2)]
            public short type;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBMDL_3
        {
            public short datano;
            public short type;
            public FocasLibBase.MODAL_AUX_data aux = new FocasLibBase.MODAL_AUX_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBMDL_4
        {
            public short datano;
            public short type;
            public FocasLibBase.MODAL_RAUX1_data raux1 = new FocasLibBase.MODAL_RAUX1_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBMDL_5
        {
            public short datano;
            public short type;
            public FocasLibBase.MODAL_RAUX2_data raux2 = new FocasLibBase.MODAL_RAUX2_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBMDLC
        {
            public short from;
            public short dram;
            public short sram;
            public short pmc;
            public short crtc;
            public short servo12;
            public short servo34;
            public short servo56;
            public short servo78;
            public short sic;
            public short pos_lsi;
            public short hi_aio;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=12)]
            public short[] reserve;
            public short drmmrc;
            public short drmarc;
            public short pmcmrc;
            public short dmaarc;
            public short iopt;
            public short hdiio;
            public short gm2gr1;
            public short crtgr2;
            public short gm1gr2;
            public short gm2gr2;
            public short cmmrb;
            public short sv5axs;
            public short sv7axs;
            public short sicaxs;
            public short posaxs;
            public short hamaxs;
            public short romr64;
            public short srmr64;
            public short dr1r64;
            public short dr2r64;
            public short iopio2;
            public short hdiio2;
            public short cmmrb2;
            public short romfap;
            public short srmfap;
            public short drmfap;
            public short drmare;
            public short pmcmre;
            public short dmaare;
            public short frmbgg;
            public short drmbgg;
            public short asrbgg;
            public short edtpsc;
            public short slcpsc;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x22)]
            public short[] reserve2;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBMGRP
        {
            public FocasLibBase.ODBMGRP_data mgrp1 = new FocasLibBase.ODBMGRP_data();
            public FocasLibBase.ODBMGRP_data mgrp2 = new FocasLibBase.ODBMGRP_data();
            public FocasLibBase.ODBMGRP_data mgrp3 = new FocasLibBase.ODBMGRP_data();
            public FocasLibBase.ODBMGRP_data mgrp4 = new FocasLibBase.ODBMGRP_data();
            public FocasLibBase.ODBMGRP_data mgrp5 = new FocasLibBase.ODBMGRP_data();
            public FocasLibBase.ODBMGRP_data mgrp6 = new FocasLibBase.ODBMGRP_data();
            public FocasLibBase.ODBMGRP_data mgrp7 = new FocasLibBase.ODBMGRP_data();
            public FocasLibBase.ODBMGRP_data mgrp8 = new FocasLibBase.ODBMGRP_data();
            public FocasLibBase.ODBMGRP_data mgrp9 = new FocasLibBase.ODBMGRP_data();
            public FocasLibBase.ODBMGRP_data mgrp10 = new FocasLibBase.ODBMGRP_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBMGRP_data
        {
            public int m_code;
            public short grp_no;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x15)]
            public string m_name = new string(' ', 0x15);
            public byte dummy;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBMIR
        {
            public short mode;
            public int mir_flag;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] mir_pos;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBMVINF
        {
            public short use_no1;
            public short use_no2;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBNC_1
        {
            public short reg_prg;
            public short unreg_prg;
            public int used_mem;
            public int unused_mem;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBNC_2
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x1f)]
            public char[] asc;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBNODE
        {
            public int node_no;
            public int io_base;
            public int status;
            public int cnc_type;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=20)]
            public string node_name = new string(' ', 20);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBOMHIS
        {
            public FocasLibBase.ODBOMHIS_data omhis1 = new FocasLibBase.ODBOMHIS_data();
            public FocasLibBase.ODBOMHIS_data omhis2 = new FocasLibBase.ODBOMHIS_data();
            public FocasLibBase.ODBOMHIS_data omhis3 = new FocasLibBase.ODBOMHIS_data();
            public FocasLibBase.ODBOMHIS_data omhis4 = new FocasLibBase.ODBOMHIS_data();
            public FocasLibBase.ODBOMHIS_data omhis5 = new FocasLibBase.ODBOMHIS_data();
            public FocasLibBase.ODBOMHIS_data omhis6 = new FocasLibBase.ODBOMHIS_data();
            public FocasLibBase.ODBOMHIS_data omhis7 = new FocasLibBase.ODBOMHIS_data();
            public FocasLibBase.ODBOMHIS_data omhis8 = new FocasLibBase.ODBOMHIS_data();
            public FocasLibBase.ODBOMHIS_data omhis9 = new FocasLibBase.ODBOMHIS_data();
            public FocasLibBase.ODBOMHIS_data omhis10 = new FocasLibBase.ODBOMHIS_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBOMHIS_data
        {
            public short om_no;
            public short year;
            public short month;
            public short day;
            public short hour;
            public short minute;
            public short second;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x100)]
            public string om_msg = new string(' ', 0x100);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBOMHIS2
        {
            public ushort s_no;
            public ushort e_no;
            public FocasLibBase.OPM_HIS opm_his = new FocasLibBase.OPM_HIS();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBOMHIS2_data
        {
            public short dsp_flg;
            public short om_no;
            public short year;
            public short month;
            public short day;
            public short hour;
            public short minute;
            public short second;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x100)]
            public string alm_msg = new string(' ', 0x100);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBOMIF
        {
            public ushort om_max;
            public ushort om_sum;
            public ushort om_char;
        }

        [StructLayout(LayoutKind.Explicit)]
        public class ODBOPHIS
        {
            [FieldOffset(0)]
            public FocasLibBase.REC_ALM2_data rec_alm = new FocasLibBase.REC_ALM2_data();
            [FieldOffset(0)]
            public FocasLibBase.REC_DATE2_data rec_date = new FocasLibBase.REC_DATE2_data();
            [FieldOffset(0)]
            public FocasLibBase.REC_MDI2_data rec_mdi = new FocasLibBase.REC_MDI2_data();
            [FieldOffset(0)]
            public FocasLibBase.REC_SGN2_data rec_sgn = new FocasLibBase.REC_SGN2_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBOPHIS4_1
        {
            public FocasLibBase.REC_MDI4_data rec_mdi1 = new FocasLibBase.REC_MDI4_data();
            public FocasLibBase.REC_MDI4_data rec_mdi2 = new FocasLibBase.REC_MDI4_data();
            public FocasLibBase.REC_MDI4_data rec_mdi3 = new FocasLibBase.REC_MDI4_data();
            public FocasLibBase.REC_MDI4_data rec_mdi4 = new FocasLibBase.REC_MDI4_data();
            public FocasLibBase.REC_MDI4_data rec_mdi5 = new FocasLibBase.REC_MDI4_data();
            public FocasLibBase.REC_MDI4_data rec_mdi6 = new FocasLibBase.REC_MDI4_data();
            public FocasLibBase.REC_MDI4_data rec_mdi7 = new FocasLibBase.REC_MDI4_data();
            public FocasLibBase.REC_MDI4_data rec_mdi8 = new FocasLibBase.REC_MDI4_data();
            public FocasLibBase.REC_MDI4_data rec_mdi9 = new FocasLibBase.REC_MDI4_data();
            public FocasLibBase.REC_MDI4_data rec_mdi10 = new FocasLibBase.REC_MDI4_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBOPHIS4_10
        {
            public FocasLibBase.REC_WOF4_data rec_wof1 = new FocasLibBase.REC_WOF4_data();
            public FocasLibBase.REC_WOF4_data rec_wof2 = new FocasLibBase.REC_WOF4_data();
            public FocasLibBase.REC_WOF4_data rec_wof3 = new FocasLibBase.REC_WOF4_data();
            public FocasLibBase.REC_WOF4_data rec_wof4 = new FocasLibBase.REC_WOF4_data();
            public FocasLibBase.REC_WOF4_data rec_wof5 = new FocasLibBase.REC_WOF4_data();
            public FocasLibBase.REC_WOF4_data rec_wof6 = new FocasLibBase.REC_WOF4_data();
            public FocasLibBase.REC_WOF4_data rec_wof7 = new FocasLibBase.REC_WOF4_data();
            public FocasLibBase.REC_WOF4_data rec_wof8 = new FocasLibBase.REC_WOF4_data();
            public FocasLibBase.REC_WOF4_data rec_wof9 = new FocasLibBase.REC_WOF4_data();
            public FocasLibBase.REC_WOF4_data rec_wof10 = new FocasLibBase.REC_WOF4_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBOPHIS4_11
        {
            public FocasLibBase.REC_MAC4_data rec_mac1 = new FocasLibBase.REC_MAC4_data();
            public FocasLibBase.REC_MAC4_data rec_mac2 = new FocasLibBase.REC_MAC4_data();
            public FocasLibBase.REC_MAC4_data rec_mac3 = new FocasLibBase.REC_MAC4_data();
            public FocasLibBase.REC_MAC4_data rec_mac4 = new FocasLibBase.REC_MAC4_data();
            public FocasLibBase.REC_MAC4_data rec_mac5 = new FocasLibBase.REC_MAC4_data();
            public FocasLibBase.REC_MAC4_data rec_mac6 = new FocasLibBase.REC_MAC4_data();
            public FocasLibBase.REC_MAC4_data rec_mac7 = new FocasLibBase.REC_MAC4_data();
            public FocasLibBase.REC_MAC4_data rec_mac8 = new FocasLibBase.REC_MAC4_data();
            public FocasLibBase.REC_MAC4_data rec_mac9 = new FocasLibBase.REC_MAC4_data();
            public FocasLibBase.REC_MAC4_data rec_mac10 = new FocasLibBase.REC_MAC4_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBOPHIS4_2
        {
            public FocasLibBase.REC_SGN4_data rec_sgn1 = new FocasLibBase.REC_SGN4_data();
            public FocasLibBase.REC_SGN4_data rec_sgn2 = new FocasLibBase.REC_SGN4_data();
            public FocasLibBase.REC_SGN4_data rec_sgn3 = new FocasLibBase.REC_SGN4_data();
            public FocasLibBase.REC_SGN4_data rec_sgn4 = new FocasLibBase.REC_SGN4_data();
            public FocasLibBase.REC_SGN4_data rec_sgn5 = new FocasLibBase.REC_SGN4_data();
            public FocasLibBase.REC_SGN4_data rec_sgn6 = new FocasLibBase.REC_SGN4_data();
            public FocasLibBase.REC_SGN4_data rec_sgn7 = new FocasLibBase.REC_SGN4_data();
            public FocasLibBase.REC_SGN4_data rec_sgn8 = new FocasLibBase.REC_SGN4_data();
            public FocasLibBase.REC_SGN4_data rec_sgn9 = new FocasLibBase.REC_SGN4_data();
            public FocasLibBase.REC_SGN4_data rec_sgn10 = new FocasLibBase.REC_SGN4_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBOPHIS4_3
        {
            public FocasLibBase.REC_ALM4_data rec_alm1 = new FocasLibBase.REC_ALM4_data();
            public FocasLibBase.REC_ALM4_data rec_alm2 = new FocasLibBase.REC_ALM4_data();
            public FocasLibBase.REC_ALM4_data rec_alm3 = new FocasLibBase.REC_ALM4_data();
            public FocasLibBase.REC_ALM4_data rec_alm4 = new FocasLibBase.REC_ALM4_data();
            public FocasLibBase.REC_ALM4_data rec_alm5 = new FocasLibBase.REC_ALM4_data();
            public FocasLibBase.REC_ALM4_data rec_alm6 = new FocasLibBase.REC_ALM4_data();
            public FocasLibBase.REC_ALM4_data rec_alm7 = new FocasLibBase.REC_ALM4_data();
            public FocasLibBase.REC_ALM4_data rec_alm8 = new FocasLibBase.REC_ALM4_data();
            public FocasLibBase.REC_ALM4_data rec_alm9 = new FocasLibBase.REC_ALM4_data();
            public FocasLibBase.REC_ALM4_data rec_alm10 = new FocasLibBase.REC_ALM4_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBOPHIS4_4
        {
            public FocasLibBase.REC_DATE4_data rec_date1 = new FocasLibBase.REC_DATE4_data();
            public FocasLibBase.REC_DATE4_data rec_date2 = new FocasLibBase.REC_DATE4_data();
            public FocasLibBase.REC_DATE4_data rec_date3 = new FocasLibBase.REC_DATE4_data();
            public FocasLibBase.REC_DATE4_data rec_date4 = new FocasLibBase.REC_DATE4_data();
            public FocasLibBase.REC_DATE4_data rec_date5 = new FocasLibBase.REC_DATE4_data();
            public FocasLibBase.REC_DATE4_data rec_date6 = new FocasLibBase.REC_DATE4_data();
            public FocasLibBase.REC_DATE4_data rec_date7 = new FocasLibBase.REC_DATE4_data();
            public FocasLibBase.REC_DATE4_data rec_date8 = new FocasLibBase.REC_DATE4_data();
            public FocasLibBase.REC_DATE4_data rec_date9 = new FocasLibBase.REC_DATE4_data();
            public FocasLibBase.REC_DATE4_data rec_date10 = new FocasLibBase.REC_DATE4_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBOPHIS4_5
        {
            public FocasLibBase.REC_IAL4_data rec_ial1 = new FocasLibBase.REC_IAL4_data();
            public FocasLibBase.REC_IAL4_data rec_ial2 = new FocasLibBase.REC_IAL4_data();
            public FocasLibBase.REC_IAL4_data rec_ial3 = new FocasLibBase.REC_IAL4_data();
            public FocasLibBase.REC_IAL4_data rec_ial4 = new FocasLibBase.REC_IAL4_data();
            public FocasLibBase.REC_IAL4_data rec_ial5 = new FocasLibBase.REC_IAL4_data();
            public FocasLibBase.REC_IAL4_data rec_ial6 = new FocasLibBase.REC_IAL4_data();
            public FocasLibBase.REC_IAL4_data rec_ial7 = new FocasLibBase.REC_IAL4_data();
            public FocasLibBase.REC_IAL4_data rec_ial8 = new FocasLibBase.REC_IAL4_data();
            public FocasLibBase.REC_IAL4_data rec_ial9 = new FocasLibBase.REC_IAL4_data();
            public FocasLibBase.REC_IAL4_data rec_ial10 = new FocasLibBase.REC_IAL4_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBOPHIS4_6
        {
            public FocasLibBase.REC_MAL4_data rec_mal1 = new FocasLibBase.REC_MAL4_data();
            public FocasLibBase.REC_MAL4_data rec_mal2 = new FocasLibBase.REC_MAL4_data();
            public FocasLibBase.REC_MAL4_data rec_mal3 = new FocasLibBase.REC_MAL4_data();
            public FocasLibBase.REC_MAL4_data rec_mal4 = new FocasLibBase.REC_MAL4_data();
            public FocasLibBase.REC_MAL4_data rec_mal5 = new FocasLibBase.REC_MAL4_data();
            public FocasLibBase.REC_MAL4_data rec_mal6 = new FocasLibBase.REC_MAL4_data();
            public FocasLibBase.REC_MAL4_data rec_mal7 = new FocasLibBase.REC_MAL4_data();
            public FocasLibBase.REC_MAL4_data rec_mal8 = new FocasLibBase.REC_MAL4_data();
            public FocasLibBase.REC_MAL4_data rec_mal9 = new FocasLibBase.REC_MAL4_data();
            public FocasLibBase.REC_MAL4_data rec_mal10 = new FocasLibBase.REC_MAL4_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBOPHIS4_7
        {
            public FocasLibBase.REC_OPM4_data rec_opm1 = new FocasLibBase.REC_OPM4_data();
            public FocasLibBase.REC_OPM4_data rec_opm2 = new FocasLibBase.REC_OPM4_data();
            public FocasLibBase.REC_OPM4_data rec_opm3 = new FocasLibBase.REC_OPM4_data();
            public FocasLibBase.REC_OPM4_data rec_opm4 = new FocasLibBase.REC_OPM4_data();
            public FocasLibBase.REC_OPM4_data rec_opm5 = new FocasLibBase.REC_OPM4_data();
            public FocasLibBase.REC_OPM4_data rec_opm6 = new FocasLibBase.REC_OPM4_data();
            public FocasLibBase.REC_OPM4_data rec_opm7 = new FocasLibBase.REC_OPM4_data();
            public FocasLibBase.REC_OPM4_data rec_opm8 = new FocasLibBase.REC_OPM4_data();
            public FocasLibBase.REC_OPM4_data rec_opm9 = new FocasLibBase.REC_OPM4_data();
            public FocasLibBase.REC_OPM4_data rec_opm10 = new FocasLibBase.REC_OPM4_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBOPHIS4_8
        {
            public FocasLibBase.REC_OFS4_data rec_ofs1 = new FocasLibBase.REC_OFS4_data();
            public FocasLibBase.REC_OFS4_data rec_ofs2 = new FocasLibBase.REC_OFS4_data();
            public FocasLibBase.REC_OFS4_data rec_ofs3 = new FocasLibBase.REC_OFS4_data();
            public FocasLibBase.REC_OFS4_data rec_ofs4 = new FocasLibBase.REC_OFS4_data();
            public FocasLibBase.REC_OFS4_data rec_ofs5 = new FocasLibBase.REC_OFS4_data();
            public FocasLibBase.REC_OFS4_data rec_ofs6 = new FocasLibBase.REC_OFS4_data();
            public FocasLibBase.REC_OFS4_data rec_ofs7 = new FocasLibBase.REC_OFS4_data();
            public FocasLibBase.REC_OFS4_data rec_ofs8 = new FocasLibBase.REC_OFS4_data();
            public FocasLibBase.REC_OFS4_data rec_ofs9 = new FocasLibBase.REC_OFS4_data();
            public FocasLibBase.REC_OFS4_data rec_ofs10 = new FocasLibBase.REC_OFS4_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBOPHIS4_9
        {
            public FocasLibBase.REC_PRM4_data rec_prm1 = new FocasLibBase.REC_PRM4_data();
            public FocasLibBase.REC_PRM4_data rec_prm2 = new FocasLibBase.REC_PRM4_data();
            public FocasLibBase.REC_PRM4_data rec_prm3 = new FocasLibBase.REC_PRM4_data();
            public FocasLibBase.REC_PRM4_data rec_prm4 = new FocasLibBase.REC_PRM4_data();
            public FocasLibBase.REC_PRM4_data rec_prm5 = new FocasLibBase.REC_PRM4_data();
            public FocasLibBase.REC_PRM4_data rec_prm6 = new FocasLibBase.REC_PRM4_data();
            public FocasLibBase.REC_PRM4_data rec_prm7 = new FocasLibBase.REC_PRM4_data();
            public FocasLibBase.REC_PRM4_data rec_prm8 = new FocasLibBase.REC_PRM4_data();
            public FocasLibBase.REC_PRM4_data rec_prm9 = new FocasLibBase.REC_PRM4_data();
            public FocasLibBase.REC_PRM4_data rec_prm10 = new FocasLibBase.REC_PRM4_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPARAIF
        {
            public ushort info_no;
            public short prev_no;
            public short next_no;
            public FocasLibBase.ODBPARAIF1 info = new FocasLibBase.ODBPARAIF1();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPARAIF_info
        {
            public short prm_no;
            public short prm_type;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPARAIF1
        {
            public FocasLibBase.ODBPARAIF_info info1 = new FocasLibBase.ODBPARAIF_info();
            public FocasLibBase.ODBPARAIF_info info2 = new FocasLibBase.ODBPARAIF_info();
            public FocasLibBase.ODBPARAIF_info info3 = new FocasLibBase.ODBPARAIF_info();
            public FocasLibBase.ODBPARAIF_info info4 = new FocasLibBase.ODBPARAIF_info();
            public FocasLibBase.ODBPARAIF_info info5 = new FocasLibBase.ODBPARAIF_info();
            public FocasLibBase.ODBPARAIF_info info6 = new FocasLibBase.ODBPARAIF_info();
            public FocasLibBase.ODBPARAIF_info info7 = new FocasLibBase.ODBPARAIF_info();
            public FocasLibBase.ODBPARAIF_info info8 = new FocasLibBase.ODBPARAIF_info();
            public FocasLibBase.ODBPARAIF_info info9 = new FocasLibBase.ODBPARAIF_info();
            public FocasLibBase.ODBPARAIF_info info10 = new FocasLibBase.ODBPARAIF_info();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPARANUM
        {
            public ushort para_min;
            public ushort para_max;
            public ushort total_no;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPDFADIR
        {
            public short data_kind;
            public short year;
            public short mon;
            public short day;
            public short hour;
            public short min;
            public short sec;
            public short dummy;
            public int dummy2;
            public int size;
            public int attr;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x24)]
            public string d_f = new string(' ', 0x24);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x34)]
            public string comment = new string(' ', 0x34);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=12)]
            public string o_time = new string(' ', 12);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPDFDRV
        {
            public short max_num;
            public short dummy;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=12)]
            public string drive1 = new string(' ', 12);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=12)]
            public string drive2 = new string(' ', 12);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=12)]
            public string drive3 = new string(' ', 12);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=12)]
            public string drive4 = new string(' ', 12);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=12)]
            public string drive5 = new string(' ', 12);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=12)]
            public string drive6 = new string(' ', 12);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=12)]
            public string drive7 = new string(' ', 12);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=12)]
            public string drive8 = new string(' ', 12);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=12)]
            public string drive9 = new string(' ', 12);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=12)]
            public string drive10 = new string(' ', 12);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=12)]
            public string drive11 = new string(' ', 12);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=12)]
            public string drive12 = new string(' ', 12);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=12)]
            public string drive13 = new string(' ', 12);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=12)]
            public string drive14 = new string(' ', 12);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=12)]
            public string drive15 = new string(' ', 12);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=12)]
            public string drive16 = new string(' ', 12);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPDFINF
        {
            public int used_page;
            public int all_page;
            public int used_dir;
            public int all_dir;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPDFNFIL
        {
            public short dir_num;
            public short file_num;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPDFSDIR
        {
            public short sub_exist;
            public short dummy;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x24)]
            public string d_f = new string(' ', 0x24);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPM
        {
            public int datano;
            public short dummy;
            public int mcr_val;
            public short dec_val;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPMCADR
        {
            public uint io_adr;
            public short datano;
            public FocasLibBase.ODBPMCADR1 info = new FocasLibBase.ODBPMCADR1();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPMCADR_info
        {
            public byte pmc_adr;
            public byte adr_attr;
            public ushort offset;
            public ushort top;
            public ushort num;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPMCADR1
        {
            public FocasLibBase.ODBPMCADR_info info1 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info2 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info3 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info4 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info5 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info6 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info7 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info8 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info9 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info10 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info11 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info12 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info13 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info14 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info15 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info16 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info17 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info18 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info19 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info20 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info21 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info22 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info23 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info24 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info25 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info26 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info27 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info28 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info29 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info30 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info31 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info32 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info33 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info34 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info35 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info36 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info37 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info38 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info39 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info40 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info41 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info42 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info43 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info44 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info45 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info46 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info47 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info48 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info49 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info50 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info51 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info52 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info53 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info54 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info55 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info56 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info57 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info58 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info59 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info60 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info61 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info62 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info63 = new FocasLibBase.ODBPMCADR_info();
            public FocasLibBase.ODBPMCADR_info info64 = new FocasLibBase.ODBPMCADR_info();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPMCALM
        {
            public FocasLibBase.ODBPMCALM_data msg1 = new FocasLibBase.ODBPMCALM_data();
            public FocasLibBase.ODBPMCALM_data msg2 = new FocasLibBase.ODBPMCALM_data();
            public FocasLibBase.ODBPMCALM_data msg3 = new FocasLibBase.ODBPMCALM_data();
            public FocasLibBase.ODBPMCALM_data msg4 = new FocasLibBase.ODBPMCALM_data();
            public FocasLibBase.ODBPMCALM_data msg5 = new FocasLibBase.ODBPMCALM_data();
            public FocasLibBase.ODBPMCALM_data msg6 = new FocasLibBase.ODBPMCALM_data();
            public FocasLibBase.ODBPMCALM_data msg7 = new FocasLibBase.ODBPMCALM_data();
            public FocasLibBase.ODBPMCALM_data msg8 = new FocasLibBase.ODBPMCALM_data();
            public FocasLibBase.ODBPMCALM_data msg9 = new FocasLibBase.ODBPMCALM_data();
            public FocasLibBase.ODBPMCALM_data msg10 = new FocasLibBase.ODBPMCALM_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPMCALM_data
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x80)]
            public string almmsg = new string(' ', 0x80);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPMCERR
        {
            public short err_no;
            public short err_dtno;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPMCINF
        {
            public short datano;
            public FocasLibBase.ODBPMCINF1 info = new FocasLibBase.ODBPMCINF1();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPMCINF_info
        {
            public char pmc_adr;
            public byte adr_attr;
            public ushort top_num;
            public ushort last_num;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPMCINF1
        {
            public FocasLibBase.ODBPMCINF_info info1 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info2 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info3 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info4 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info5 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info6 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info7 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info8 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info9 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info10 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info11 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info12 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info13 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info14 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info15 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info16 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info17 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info18 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info19 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info20 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info21 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info22 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info23 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info24 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info25 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info26 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info27 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info28 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info29 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info30 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info31 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info32 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info33 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info34 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info35 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info36 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info37 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info38 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info39 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info40 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info41 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info42 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info43 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info44 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info45 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info46 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info47 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info48 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info49 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info50 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info51 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info52 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info53 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info54 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info55 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info56 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info57 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info58 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info59 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info60 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info61 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info62 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info63 = new FocasLibBase.ODBPMCINF_info();
            public FocasLibBase.ODBPMCINF_info info64 = new FocasLibBase.ODBPMCINF_info();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPMCTITLE
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x30)]
            public string mtb = new string(' ', 0x30);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x30)]
            public string machine = new string(' ', 0x30);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x30)]
            public string type = new string(' ', 0x30);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=8)]
            public string prgno = new string(' ', 8);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=4)]
            public string prgvers = new string(' ', 4);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x30)]
            public string prgdraw = new string(' ', 0x30);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x20)]
            public string date = new string(' ', 0x20);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x30)]
            public string design = new string(' ', 0x30);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x30)]
            public string written = new string(' ', 0x30);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x30)]
            public string remarks = new string(' ', 0x30);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPMINF
        {
            public short use_no1;
            public short use_no2;
            public short v2_type;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPOFS
        {
            public short mode;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] ofsvct;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPOS
        {
            public FocasLibBase.POSELMALL p1 = new FocasLibBase.POSELMALL();
            public FocasLibBase.POSELMALL p2 = new FocasLibBase.POSELMALL();
            public FocasLibBase.POSELMALL p3 = new FocasLibBase.POSELMALL();
            public FocasLibBase.POSELMALL p4 = new FocasLibBase.POSELMALL();
            public FocasLibBase.POSELMALL p5 = new FocasLibBase.POSELMALL();
            public FocasLibBase.POSELMALL p6 = new FocasLibBase.POSELMALL();
            public FocasLibBase.POSELMALL p7 = new FocasLibBase.POSELMALL();
            public FocasLibBase.POSELMALL p8 = new FocasLibBase.POSELMALL();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPRFCNF
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string master_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=3)]
            public string master_ver = new string(' ', 3);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string slave_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=3)]
            public string slave_ver = new string(' ', 3);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string cntl_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=3)]
            public string cntl_ver = new string(' ', 3);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPRO
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=2)]
            public short[] dummy;
            public short data;
            public short mdata;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPRS
        {
            public short datano;
            public short type;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=5)]
            public short[] data_info;
            public int rstr_bc;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x23)]
            public int[] rstr_m;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=2)]
            public int[] rstr_t;
            public int rstr_s;
            public int rstr_b;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] dest;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] dist;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPSER
        {
            public int poserr1;
            public int poserr2;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPTIME
        {
            public short num;
            public FocasLibBase.ODBPTIME1 data = new FocasLibBase.ODBPTIME1();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPTIME_data
        {
            public int prg_no;
            public short hour;
            public byte minute;
            public byte second;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPTIME1
        {
            public FocasLibBase.ODBPTIME_data data1 = new FocasLibBase.ODBPTIME_data();
            public FocasLibBase.ODBPTIME_data data2 = new FocasLibBase.ODBPTIME_data();
            public FocasLibBase.ODBPTIME_data data3 = new FocasLibBase.ODBPTIME_data();
            public FocasLibBase.ODBPTIME_data data4 = new FocasLibBase.ODBPTIME_data();
            public FocasLibBase.ODBPTIME_data data5 = new FocasLibBase.ODBPTIME_data();
            public FocasLibBase.ODBPTIME_data data6 = new FocasLibBase.ODBPTIME_data();
            public FocasLibBase.ODBPTIME_data data7 = new FocasLibBase.ODBPTIME_data();
            public FocasLibBase.ODBPTIME_data data8 = new FocasLibBase.ODBPTIME_data();
            public FocasLibBase.ODBPTIME_data data9 = new FocasLibBase.ODBPTIME_data();
            public FocasLibBase.ODBPTIME_data data10 = new FocasLibBase.ODBPTIME_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPTLINF
        {
            public short tld_max;
            public short mlt_max;
            public short reserve;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x10)]
            public short[] tld_size;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x10)]
            public short[] mlt_size;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x10)]
            public short[] reserves;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPWOFST
        {
            public FocasLibBase.ODBPWOFST_data data1 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data2 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data3 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data4 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data5 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data6 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data7 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data8 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data9 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data10 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data11 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data12 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data13 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data14 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data15 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data16 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data17 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data18 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data19 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data20 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data21 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data22 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data23 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data24 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data25 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data26 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data27 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data28 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data29 = new FocasLibBase.ODBPWOFST_data();
            public FocasLibBase.ODBPWOFST_data data30 = new FocasLibBase.ODBPWOFST_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBPWOFST_data
        {
            public int pwratio;
            public int rfvolt;
            public ushort year;
            public ushort month;
            public ushort day;
            public ushort hour;
            public ushort minute;
            public ushort second;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBRMTDT
        {
            public short channel;
            public short kind;
            public byte year;
            public byte month;
            public byte day;
            public byte hour;
            public byte minute;
            public byte second;
            public short t_intrvl;
            public short trg_data;
            public int ins_ptr;
            public short t_delta;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x77d)]
            public short[] data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBROFS
        {
            public short mode;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=2)]
            public short[] pln_axes;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=2)]
            public int[] ofsvct;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBROT
        {
            public short mode;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=2)]
            public short[] pln_axes;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=2)]
            public int[] center;
            public int angle;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBRSTRM
        {
            public short grp_no;
            public short mem_no;
            public FocasLibBase.M_CODE1 m_code = new FocasLibBase.M_CODE1();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSCL
        {
            public short mode;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] center;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] magnif;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSD
        {
            public IntPtr chadata;
            public IntPtr count;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSEQ
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=2)]
            public short[] dummy;
            public int data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSETIF
        {
            public ushort info_no;
            public short prev_no;
            public short next_no;
            public FocasLibBase.ODBSETIF1 info = new FocasLibBase.ODBSETIF1();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSETIF_info
        {
            public short set_no;
            public short set_type;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSETIF1
        {
            public FocasLibBase.ODBSETIF_info info1 = new FocasLibBase.ODBSETIF_info();
            public FocasLibBase.ODBSETIF_info info2 = new FocasLibBase.ODBSETIF_info();
            public FocasLibBase.ODBSETIF_info info3 = new FocasLibBase.ODBSETIF_info();
            public FocasLibBase.ODBSETIF_info info4 = new FocasLibBase.ODBSETIF_info();
            public FocasLibBase.ODBSETIF_info info5 = new FocasLibBase.ODBSETIF_info();
            public FocasLibBase.ODBSETIF_info info6 = new FocasLibBase.ODBSETIF_info();
            public FocasLibBase.ODBSETIF_info info7 = new FocasLibBase.ODBSETIF_info();
            public FocasLibBase.ODBSETIF_info info8 = new FocasLibBase.ODBSETIF_info();
            public FocasLibBase.ODBSETIF_info info9 = new FocasLibBase.ODBSETIF_info();
            public FocasLibBase.ODBSETIF_info info10 = new FocasLibBase.ODBSETIF_info();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSETNUM
        {
            public ushort set_min;
            public ushort set_max;
            public ushort total_no;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSINFO
        {
            public int sramnum;
            public FocasLibBase.ODBSINFO1 info = new FocasLibBase.ODBSINFO1();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSINFO_info
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=12)]
            public string sramname = new string(' ', 12);
            public int sramsize;
            public short divnumber;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x10)]
            public string fname1 = new string(' ', 0x10);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x10)]
            public string fname2 = new string(' ', 0x10);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x10)]
            public string fname3 = new string(' ', 0x10);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x10)]
            public string fname4 = new string(' ', 0x10);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x10)]
            public string fname5 = new string(' ', 0x10);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x10)]
            public string fname6 = new string(' ', 0x10);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSINFO1
        {
            public FocasLibBase.ODBSINFO_info info1 = new FocasLibBase.ODBSINFO_info();
            public FocasLibBase.ODBSINFO_info info2 = new FocasLibBase.ODBSINFO_info();
            public FocasLibBase.ODBSINFO_info info3 = new FocasLibBase.ODBSINFO_info();
            public FocasLibBase.ODBSINFO_info info4 = new FocasLibBase.ODBSINFO_info();
            public FocasLibBase.ODBSINFO_info info5 = new FocasLibBase.ODBSINFO_info();
            public FocasLibBase.ODBSINFO_info info6 = new FocasLibBase.ODBSINFO_info();
            public FocasLibBase.ODBSINFO_info info7 = new FocasLibBase.ODBSINFO_info();
            public FocasLibBase.ODBSINFO_info info8 = new FocasLibBase.ODBSINFO_info();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSLVST
        {
            public byte cnfg_stat;
            public byte prm_stat;
            public sbyte wdg_stat;
            public byte live_stat;
            public short ident_no;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSPDI
        {
            public FocasLibBase.ODBSPDI_data di1 = new FocasLibBase.ODBSPDI_data();
            public FocasLibBase.ODBSPDI_data di2 = new FocasLibBase.ODBSPDI_data();
            public FocasLibBase.ODBSPDI_data di3 = new FocasLibBase.ODBSPDI_data();
            public FocasLibBase.ODBSPDI_data di4 = new FocasLibBase.ODBSPDI_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSPDI_data
        {
            public byte sgnl1;
            public byte sgnl2;
            public byte sgnl3;
            public byte sgnl4;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSPDLNAME
        {
            public FocasLibBase.ODBSPDLNAME_data data1 = new FocasLibBase.ODBSPDLNAME_data();
            public FocasLibBase.ODBSPDLNAME_data data2 = new FocasLibBase.ODBSPDLNAME_data();
            public FocasLibBase.ODBSPDLNAME_data data3 = new FocasLibBase.ODBSPDLNAME_data();
            public FocasLibBase.ODBSPDLNAME_data data4 = new FocasLibBase.ODBSPDLNAME_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSPDLNAME_data
        {
            public byte name;
            public byte suff1;
            public byte suff2;
            public byte suff3;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSPDO
        {
            public FocasLibBase.ODBSPDO_data do1 = new FocasLibBase.ODBSPDO_data();
            public FocasLibBase.ODBSPDO_data do2 = new FocasLibBase.ODBSPDO_data();
            public FocasLibBase.ODBSPDO_data do3 = new FocasLibBase.ODBSPDO_data();
            public FocasLibBase.ODBSPDO_data do4 = new FocasLibBase.ODBSPDO_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSPDO_data
        {
            public byte sgnl1;
            public byte sgnl2;
            public byte sgnl3;
            public byte sgnl4;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSPEED
        {
            public FocasLibBase.SPEEDELM actf = new FocasLibBase.SPEEDELM();
            public FocasLibBase.SPEEDELM acts = new FocasLibBase.SPEEDELM();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSPLOAD
        {
            public FocasLibBase.ODBSPLOAD_data spload1 = new FocasLibBase.ODBSPLOAD_data();
            public FocasLibBase.ODBSPLOAD_data spload2 = new FocasLibBase.ODBSPLOAD_data();
            public FocasLibBase.ODBSPLOAD_data spload3 = new FocasLibBase.ODBSPLOAD_data();
            public FocasLibBase.ODBSPLOAD_data spload4 = new FocasLibBase.ODBSPLOAD_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSPLOAD_data
        {
            public FocasLibBase.LOADELM spload = new FocasLibBase.LOADELM();
            public FocasLibBase.LOADELM spspeed = new FocasLibBase.LOADELM();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSPN
        {
            public short datano;
            public short type;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=4)]
            public short[] data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSRCSLYT
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=4)]
            public short[] spndl;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public short[] servo;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=8)]
            public string axis_name = new string(' ', 8);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSRCSST
        {
            public short acc_element;
            public short err_general;
            public short err_id_no;
            public short err_attr;
            public short err_op_data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBST
        {
            public short dummy;
            public short tmmode;
            public short aut;
            public short run;
            public short motion;
            public short mstb;
            public short emergency;
            public short alarm;
            public short edit;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSVLOAD
        {
            public FocasLibBase.LOADELM svload1 = new FocasLibBase.LOADELM();
            public FocasLibBase.LOADELM svload2 = new FocasLibBase.LOADELM();
            public FocasLibBase.LOADELM svload3 = new FocasLibBase.LOADELM();
            public FocasLibBase.LOADELM svload4 = new FocasLibBase.LOADELM();
            public FocasLibBase.LOADELM svload5 = new FocasLibBase.LOADELM();
            public FocasLibBase.LOADELM svload6 = new FocasLibBase.LOADELM();
            public FocasLibBase.LOADELM svload7 = new FocasLibBase.LOADELM();
            public FocasLibBase.LOADELM svload8 = new FocasLibBase.LOADELM();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSYS
        {
            public short addinfo;
            public short max_axis;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=2)]
            public char[] cnc_type;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=2)]
            public char[] mt_type;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=4)]
            public char[] series;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=4)]
            public char[] version;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=2)]
            public char[] axes;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSYSC
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x10)]
            public byte[] slot_no_p;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x10)]
            public byte[] slot_no_l;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x10)]
            public short[] mod_id;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x10)]
            public short[] soft_id;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_series1 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_series2 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_series3 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_series4 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_series5 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_series6 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_series7 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_series8 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_series9 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_series10 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_series11 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_series12 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_series13 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_series14 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_series15 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_series16 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_version1 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_version2 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_version3 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_version4 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_version5 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_version6 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_version7 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_version8 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_version9 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_version10 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_version11 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_version12 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_version13 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_version14 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_version15 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string s_version16 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x10)]
            public byte[] dummy;
            public short m_rom;
            public short s_rom;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public char[] svo_soft;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=6)]
            public char[] pmc_soft;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=6)]
            public char[] lad_soft;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public char[] mcr_soft;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=6)]
            public char[] spl1_soft;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=6)]
            public char[] spl2_soft;
            public short frmmin;
            public short drmmin;
            public short srmmin;
            public short pmcmin;
            public short crtmin;
            public short sv1min;
            public short sv3min;
            public short sicmin;
            public short posmin;
            public short drmmrc;
            public short drmarc;
            public short pmcmrc;
            public short dmaarc;
            public short iopt;
            public short hdiio;
            public short frmsub;
            public short drmsub;
            public short srmsub;
            public short sv5sub;
            public short sv7sub;
            public short sicsub;
            public short possub;
            public short hamsub;
            public short gm2gr1;
            public short crtgr2;
            public short gm1gr2;
            public short gm2gr2;
            public short cmmrb;
            public short sv5axs;
            public short sv7axs;
            public short sicaxs;
            public short posaxs;
            public short hanaxs;
            public short romr64;
            public short srmr64;
            public short dr1r64;
            public short dr2r64;
            public short iopio2;
            public short hdiio2;
            public short cmmrb2;
            public short romfap;
            public short srmfap;
            public short drmfap;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSYSEX
        {
            public short max_axis;
            public short max_spdl;
            public short max_path;
            public short max_mchn;
            public short ctrl_axis;
            public short ctrl_srvo;
            public short ctrl_spdl;
            public short ctrl_path;
            public short ctrl_mchn;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=3)]
            public short[] reserved = new short[3];
            public FocasLibBase.ODBSYSEX_data path = new FocasLibBase.ODBSYSEX_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSYSEX_data
        {
            public FocasLibBase.ODBSYSEX_path data1 = new FocasLibBase.ODBSYSEX_path();
            public FocasLibBase.ODBSYSEX_path data2 = new FocasLibBase.ODBSYSEX_path();
            public FocasLibBase.ODBSYSEX_path data3 = new FocasLibBase.ODBSYSEX_path();
            public FocasLibBase.ODBSYSEX_path data4 = new FocasLibBase.ODBSYSEX_path();
            public FocasLibBase.ODBSYSEX_path data5 = new FocasLibBase.ODBSYSEX_path();
            public FocasLibBase.ODBSYSEX_path data6 = new FocasLibBase.ODBSYSEX_path();
            public FocasLibBase.ODBSYSEX_path data7 = new FocasLibBase.ODBSYSEX_path();
            public FocasLibBase.ODBSYSEX_path data8 = new FocasLibBase.ODBSYSEX_path();
            public FocasLibBase.ODBSYSEX_path data9 = new FocasLibBase.ODBSYSEX_path();
            public FocasLibBase.ODBSYSEX_path data10 = new FocasLibBase.ODBSYSEX_path();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSYSEX_path
        {
            public short system;
            public short group;
            public short attrib;
            public short ctrl_axis;
            public short ctrl_srvo;
            public short ctrl_spdl;
            public short mchn_no;
            public short reserved;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSYSH
        {
            public FocasLibBase.ODBSYSH_data data1 = new FocasLibBase.ODBSYSH_data();
            public FocasLibBase.ODBSYSH_data data2 = new FocasLibBase.ODBSYSH_data();
            public FocasLibBase.ODBSYSH_data data3 = new FocasLibBase.ODBSYSH_data();
            public FocasLibBase.ODBSYSH_data data4 = new FocasLibBase.ODBSYSH_data();
            public FocasLibBase.ODBSYSH_data data5 = new FocasLibBase.ODBSYSH_data();
            public FocasLibBase.ODBSYSH_data data6 = new FocasLibBase.ODBSYSH_data();
            public FocasLibBase.ODBSYSH_data data7 = new FocasLibBase.ODBSYSH_data();
            public FocasLibBase.ODBSYSH_data data8 = new FocasLibBase.ODBSYSH_data();
            public FocasLibBase.ODBSYSH_data data9 = new FocasLibBase.ODBSYSH_data();
            public FocasLibBase.ODBSYSH_data data10 = new FocasLibBase.ODBSYSH_data();
            public FocasLibBase.ODBSYSH_data data11 = new FocasLibBase.ODBSYSH_data();
            public FocasLibBase.ODBSYSH_data data12 = new FocasLibBase.ODBSYSH_data();
            public FocasLibBase.ODBSYSH_data data13 = new FocasLibBase.ODBSYSH_data();
            public FocasLibBase.ODBSYSH_data data14 = new FocasLibBase.ODBSYSH_data();
            public FocasLibBase.ODBSYSH_data data15 = new FocasLibBase.ODBSYSH_data();
            public FocasLibBase.ODBSYSH_data data16 = new FocasLibBase.ODBSYSH_data();
            public FocasLibBase.ODBSYSH_data data17 = new FocasLibBase.ODBSYSH_data();
            public FocasLibBase.ODBSYSH_data data18 = new FocasLibBase.ODBSYSH_data();
            public FocasLibBase.ODBSYSH_data data19 = new FocasLibBase.ODBSYSH_data();
            public FocasLibBase.ODBSYSH_data data20 = new FocasLibBase.ODBSYSH_data();
            public FocasLibBase.ODBSYSH_data data21 = new FocasLibBase.ODBSYSH_data();
            public FocasLibBase.ODBSYSH_data data22 = new FocasLibBase.ODBSYSH_data();
            public FocasLibBase.ODBSYSH_data data23 = new FocasLibBase.ODBSYSH_data();
            public FocasLibBase.ODBSYSH_data data24 = new FocasLibBase.ODBSYSH_data();
            public FocasLibBase.ODBSYSH_data data25 = new FocasLibBase.ODBSYSH_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSYSH_data
        {
            public uint id1;
            public uint id2;
            public short group_id;
            public short hard_id;
            public short hard_num;
            public short slot_no;
            public short id1_format;
            public short id2_format;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSYSS
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x10)]
            public byte[] slot_no_p;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x10)]
            public byte[] slot_no_l;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x10)]
            public short[] module_id;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x10)]
            public short[] soft_id;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series1 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series2 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series3 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series4 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series5 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series6 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series7 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series8 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series9 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series10 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series11 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series12 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series13 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series14 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series15 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series16 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version1 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version2 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version3 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version4 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version5 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version6 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version7 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version8 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version9 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version10 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version11 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version12 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version13 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version14 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version15 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version16 = new string(' ', 5);
            public short soft_inst;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string boot_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string boot_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string servo_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string servo_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmc_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmc_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string ladder_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string ladder_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string mcrlib_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string mcrlib_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string mcrapl_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string mcrapl_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string spl1_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string spl1_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string spl2_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string spl2_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string spl3_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string spl3_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string c_exelib_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string c_exelib_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string c_exeapl_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string c_exeapl_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string int_vga_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string int_vga_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string out_vga_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string out_vga_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmm_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmm_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmc_mng_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmc_mng_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmc_shin_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmc_shin_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmc_shout_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmc_shout_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmc_c_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmc_c_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmc_edit_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmc_edit_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string lddr_mng_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string lddr_mng_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string lddr_apl_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string lddr_apl_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string spl4_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string spl4_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string mcr2_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string mcr2_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string mcr3_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string mcr3_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string eth_boot_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string eth_boot_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=40)]
            public byte[] reserve;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSYSS2
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x10)]
            public byte[] slot_no_p;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x10)]
            public byte[] slot_no_l;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x10)]
            public short[] module_id;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x10)]
            public short[] soft_id;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series1 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series2 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series3 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series4 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series5 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series6 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series7 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series8 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series9 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series10 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series11 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series12 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series13 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series14 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series15 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_series16 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version1 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version2 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version3 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version4 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version5 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version6 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version7 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version8 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version9 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version10 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version11 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version12 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version13 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version14 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version15 = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string soft_version16 = new string(' ', 5);
            public short soft_inst;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string boot_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string boot_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string servo_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string servo_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmc_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmc_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string ladder_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string ladder_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string mcrlib_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string mcrlib_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string mcrapl_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string mcrapl_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string spl1_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string spl1_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string spl2_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string spl2_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string spl3_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string spl3_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string c_exelib_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string c_exelib_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string c_exeapl_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string c_exeapl_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string int_vga_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string int_vga_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string out_vga_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string out_vga_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmm_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmm_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmc_mng_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmc_mng_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmc_shin_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmc_shin_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmc_shout_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmc_shout_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmc_c_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmc_c_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmc_edit_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string pmc_edit_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string lddr_mng_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string lddr_mng_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string lddr_apl_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string lddr_apl_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string spl4_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string spl4_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string mcr2_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string mcr2_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string mcr3_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string mcr3_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string eth_boot_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string eth_boot_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=40)]
            public byte[] reserve;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string embEthe_ser = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=5)]
            public string embEthe_ver = new string(' ', 5);
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=190)]
            public byte[] reserve2;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSYSS3
        {
            public FocasLibBase.ODBSYSS3_data p1 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p2 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p3 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p4 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p5 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p6 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p7 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p8 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p9 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p10 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p11 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p12 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p13 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p14 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p15 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p16 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p17 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p18 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p19 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p20 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p21 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p22 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p23 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p24 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p25 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p26 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p27 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p28 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p29 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p30 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p31 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p32 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p33 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p34 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p35 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p36 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p37 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p38 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p39 = new FocasLibBase.ODBSYSS3_data();
            public FocasLibBase.ODBSYSS3_data p40 = new FocasLibBase.ODBSYSS3_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBSYSS3_data
        {
            public short soft_id;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=5)]
            public char[] soft_series;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=5)]
            public char[] soft_edition;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBTG
        {
            public short grp_num;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=2)]
            public short[] dummy;
            public int ntool;
            public int life;
            public int count;
            public FocasLibBase.ODBTG1 data = new FocasLibBase.ODBTG1();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBTG_data
        {
            public int tuse_num;
            public int tool_num;
            public int length_num;
            public int radius_num;
            public int tinfo;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBTG1
        {
            public FocasLibBase.ODBTG_data data1 = new FocasLibBase.ODBTG_data();
            public FocasLibBase.ODBTG_data data2 = new FocasLibBase.ODBTG_data();
            public FocasLibBase.ODBTG_data data3 = new FocasLibBase.ODBTG_data();
            public FocasLibBase.ODBTG_data data4 = new FocasLibBase.ODBTG_data();
            public FocasLibBase.ODBTG_data data5 = new FocasLibBase.ODBTG_data();

            public FocasLibBase.ODBTG_data data6 = new FocasLibBase.ODBTG_data();
            public FocasLibBase.ODBTG_data data7 = new FocasLibBase.ODBTG_data();
            public FocasLibBase.ODBTG_data data8 = new FocasLibBase.ODBTG_data();
            public FocasLibBase.ODBTG_data data9 = new FocasLibBase.ODBTG_data();
            public FocasLibBase.ODBTG_data data10 = new FocasLibBase.ODBTG_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBTLIFE1
        {
            public short dummy;
            public short type;
            public int data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBTLIFE2
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=2)]
            public short[] dummy;
            public int data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBTLIFE3
        {
            public short datano;
            public short dummy;
            public int data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBTLIFE4
        {
            public short datano;
            public short type;
            public int data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBTLIFE5
        {
            public int dummy;
            public int type;
            public int data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBTLINF
        {
            public short ofs_type;
            public short use_no;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBTLINF2
        {
            public short ofs_type;
            public short use_no;
            public short ofs_enable;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBTLINFO
        {
            public int max_group;
            public int max_tool;
            public int max_minute;
            public int max_cycle;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBTLUSE
        {
            public short s_grp;
            public short dummy;
            public short e_grp;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=5)]
            public int[] data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBTOFS
        {
            public short datano;
            public short type;
            public int data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBUP
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=2)]
            public short[] dummy;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x100)]
            public char[] data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBUSEGR
        {
            public short datano;
            public short type;
            public int next;
            public int use;
            public int slct;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBUSEGRP
        {
            public int next;
            public int use;
            public int slct;
            public int opt_next;
            public int opt_use;
            public int opt_slct;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBWVDT
        {
            public short channel;
            public short kind;
            public FocasLibBase.ODBWVDT_u u = new FocasLibBase.ODBWVDT_u();
            public byte year;
            public byte month;
            public byte day;
            public byte hour;
            public byte minute;
            public byte second;
            public short t_cycle;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x2000)]
            public short[] data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBWVDT_axis
        {
            public short axis;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class ODBWVDT_io
        {
            public byte adr;
            public byte bit;
            public short no;
        }

        [StructLayout(LayoutKind.Explicit)]
        public class ODBWVDT_u
        {
            [FieldOffset(0)]
            public FocasLibBase.ODBWVDT_axis axis = new FocasLibBase.ODBWVDT_axis();
            [FieldOffset(0)]
            public FocasLibBase.ODBWVDT_io io = new FocasLibBase.ODBWVDT_io();
        }

        [StructLayout(LayoutKind.Explicit)]
        public class OFS_1
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=5), FieldOffset(0)]
            public int[] m_ofs = new int[5];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=5), FieldOffset(0)]
            public int[] m_ofs_a = new int[5];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=5), FieldOffset(0)]
            public int[] t_ofs = new int[5];
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=5), FieldOffset(0)]
            public short[] t_tip = new short[5];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class OFS_2
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=10)]
            public int[] m_ofs_b = new int[10];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class OFS_3
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=20)]
            public int[] m_ofs_c = new int[20];
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class OPM_HIS
        {
            public FocasLibBase.ODBOMHIS2_data data1 = new FocasLibBase.ODBOMHIS2_data();
            public FocasLibBase.ODBOMHIS2_data data2 = new FocasLibBase.ODBOMHIS2_data();
            public FocasLibBase.ODBOMHIS2_data data3 = new FocasLibBase.ODBOMHIS2_data();
            public FocasLibBase.ODBOMHIS2_data data4 = new FocasLibBase.ODBOMHIS2_data();
            public FocasLibBase.ODBOMHIS2_data data5 = new FocasLibBase.ODBOMHIS2_data();
            public FocasLibBase.ODBOMHIS2_data data6 = new FocasLibBase.ODBOMHIS2_data();
            public FocasLibBase.ODBOMHIS2_data data7 = new FocasLibBase.ODBOMHIS2_data();
            public FocasLibBase.ODBOMHIS2_data data8 = new FocasLibBase.ODBOMHIS2_data();
            public FocasLibBase.ODBOMHIS2_data data9 = new FocasLibBase.ODBOMHIS2_data();
            public FocasLibBase.ODBOMHIS2_data data10 = new FocasLibBase.ODBOMHIS2_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class OPMSG
        {
            public FocasLibBase.OPMSG_data msg1 = new FocasLibBase.OPMSG_data();
            public FocasLibBase.OPMSG_data msg2 = new FocasLibBase.OPMSG_data();
            public FocasLibBase.OPMSG_data msg3 = new FocasLibBase.OPMSG_data();
            public FocasLibBase.OPMSG_data msg4 = new FocasLibBase.OPMSG_data();
            public FocasLibBase.OPMSG_data msg5 = new FocasLibBase.OPMSG_data();
            public FocasLibBase.OPMSG_data msg6 = new FocasLibBase.OPMSG_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class OPMSG_data
        {
            public short datano;
            public short type;
            public short char_num;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x100)]
            public string data = new string(' ', 0x100);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class OPMSG2
        {
            public FocasLibBase.OPMSG2_data msg1 = new FocasLibBase.OPMSG2_data();
            public FocasLibBase.OPMSG2_data msg2 = new FocasLibBase.OPMSG2_data();
            public FocasLibBase.OPMSG2_data msg3 = new FocasLibBase.OPMSG2_data();
            public FocasLibBase.OPMSG2_data msg4 = new FocasLibBase.OPMSG2_data();
            public FocasLibBase.OPMSG2_data msg5 = new FocasLibBase.OPMSG2_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class OPMSG2_data
        {
            public short datano;
            public short type;
            public short char_num;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x40)]
            public string data = new string(' ', 0x40);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class OPMSG3
        {
            public FocasLibBase.OPMSG3_data msg1 = new FocasLibBase.OPMSG3_data();
            public FocasLibBase.OPMSG3_data msg2 = new FocasLibBase.OPMSG3_data();
            public FocasLibBase.OPMSG3_data msg3 = new FocasLibBase.OPMSG3_data();
            public FocasLibBase.OPMSG3_data msg4 = new FocasLibBase.OPMSG3_data();
            public FocasLibBase.OPMSG3_data msg5 = new FocasLibBase.OPMSG3_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class OPMSG3_data
        {
            public short datano;
            public short type;
            public short char_num;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x100)]
            public string data = new string(' ', 0x100);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class POSELM
        {
            public int data;
            public short dec;
            public short unit;
            public short disp;
            public char name;
            public char suff;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class POSELMALL
        {
            public FocasLibBase.POSELM abs = new FocasLibBase.POSELM();
            public FocasLibBase.POSELM mach = new FocasLibBase.POSELM();
            public FocasLibBase.POSELM rel = new FocasLibBase.POSELM();
            public FocasLibBase.POSELM dist = new FocasLibBase.POSELM();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class PRGDIR
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x100)]
            public char[] prg_data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class PRGDIR2
        {
            public FocasLibBase.PRGDIR2_data dir1 = new FocasLibBase.PRGDIR2_data();
            public FocasLibBase.PRGDIR2_data dir2 = new FocasLibBase.PRGDIR2_data();
            public FocasLibBase.PRGDIR2_data dir3 = new FocasLibBase.PRGDIR2_data();
            public FocasLibBase.PRGDIR2_data dir4 = new FocasLibBase.PRGDIR2_data();
            public FocasLibBase.PRGDIR2_data dir5 = new FocasLibBase.PRGDIR2_data();
            public FocasLibBase.PRGDIR2_data dir6 = new FocasLibBase.PRGDIR2_data();
            public FocasLibBase.PRGDIR2_data dir7 = new FocasLibBase.PRGDIR2_data();
            public FocasLibBase.PRGDIR2_data dir8 = new FocasLibBase.PRGDIR2_data();
            public FocasLibBase.PRGDIR2_data dir9 = new FocasLibBase.PRGDIR2_data();
            public FocasLibBase.PRGDIR2_data dir10 = new FocasLibBase.PRGDIR2_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class PRGDIR2_data
        {
            public short number;
            public int length;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x33)]
            public string comment = new string(' ', 0x33);
            public byte dummy;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class PRGDIR3
        {
            public FocasLibBase.PRGDIR3_data dir1 = new FocasLibBase.PRGDIR3_data();
            public FocasLibBase.PRGDIR3_data dir2 = new FocasLibBase.PRGDIR3_data();
            public FocasLibBase.PRGDIR3_data dir3 = new FocasLibBase.PRGDIR3_data();
            public FocasLibBase.PRGDIR3_data dir4 = new FocasLibBase.PRGDIR3_data();
            public FocasLibBase.PRGDIR3_data dir5 = new FocasLibBase.PRGDIR3_data();
            public FocasLibBase.PRGDIR3_data dir6 = new FocasLibBase.PRGDIR3_data();
            public FocasLibBase.PRGDIR3_data dir7 = new FocasLibBase.PRGDIR3_data();
            public FocasLibBase.PRGDIR3_data dir8 = new FocasLibBase.PRGDIR3_data();
            public FocasLibBase.PRGDIR3_data dir9 = new FocasLibBase.PRGDIR3_data();
            public FocasLibBase.PRGDIR3_data dir10 = new FocasLibBase.PRGDIR3_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class PRGDIR3_data
        {
            public int number;
            public int length;
            public int page;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x34)]
            public string comment = new string(' ', 0x34);
            public FocasLibBase.DIR3_MDATE mdate = new FocasLibBase.DIR3_MDATE();
            public FocasLibBase.DIR3_CDATE cdate = new FocasLibBase.DIR3_CDATE();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class PRGDIR4
        {
            public FocasLibBase.PRGDIR4_data dir1 = new FocasLibBase.PRGDIR4_data();
            public FocasLibBase.PRGDIR4_data dir2 = new FocasLibBase.PRGDIR4_data();
            public FocasLibBase.PRGDIR4_data dir3 = new FocasLibBase.PRGDIR4_data();
            public FocasLibBase.PRGDIR4_data dir4 = new FocasLibBase.PRGDIR4_data();
            public FocasLibBase.PRGDIR4_data dir5 = new FocasLibBase.PRGDIR4_data();
            public FocasLibBase.PRGDIR4_data dir6 = new FocasLibBase.PRGDIR4_data();
            public FocasLibBase.PRGDIR4_data dir7 = new FocasLibBase.PRGDIR4_data();
            public FocasLibBase.PRGDIR4_data dir8 = new FocasLibBase.PRGDIR4_data();
            public FocasLibBase.PRGDIR4_data dir9 = new FocasLibBase.PRGDIR4_data();
            public FocasLibBase.PRGDIR4_data dir10 = new FocasLibBase.PRGDIR4_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class PRGDIR4_data
        {
            public int number;
            public int length;
            public int page;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x34)]
            public string comment = new string(' ', 0x34);
            public FocasLibBase.DIR4_MDATE mdate = new FocasLibBase.DIR4_MDATE();
            public FocasLibBase.DIR4_CDATE cdate = new FocasLibBase.DIR4_CDATE();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class PRGDIRTM
        {
            public FocasLibBase.PRGDIRTM_data data1 = new FocasLibBase.PRGDIRTM_data();
            public FocasLibBase.PRGDIRTM_data data2 = new FocasLibBase.PRGDIRTM_data();
            public FocasLibBase.PRGDIRTM_data data3 = new FocasLibBase.PRGDIRTM_data();
            public FocasLibBase.PRGDIRTM_data data4 = new FocasLibBase.PRGDIRTM_data();
            public FocasLibBase.PRGDIRTM_data data5 = new FocasLibBase.PRGDIRTM_data();
            public FocasLibBase.PRGDIRTM_data data6 = new FocasLibBase.PRGDIRTM_data();
            public FocasLibBase.PRGDIRTM_data data7 = new FocasLibBase.PRGDIRTM_data();
            public FocasLibBase.PRGDIRTM_data data8 = new FocasLibBase.PRGDIRTM_data();
            public FocasLibBase.PRGDIRTM_data data9 = new FocasLibBase.PRGDIRTM_data();
            public FocasLibBase.PRGDIRTM_data data10 = new FocasLibBase.PRGDIRTM_data();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class PRGDIRTM_data
        {
            public int prg_no;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x33)]
            public string comment = new string(' ', 0x33);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=13)]
            public string cuttime = new string(' ', 13);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class PRGPNT
        {
            public int prog_no;
            public int blk_no;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REALDGN
        {
            public int dgn_val;
            public int dec_val;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REALDGNS
        {
            public FocasLibBase.REALDGN rdata1 = new FocasLibBase.REALDGN();
            public FocasLibBase.REALDGN rdata2 = new FocasLibBase.REALDGN();
            public FocasLibBase.REALDGN rdata3 = new FocasLibBase.REALDGN();
            public FocasLibBase.REALDGN rdata4 = new FocasLibBase.REALDGN();
            public FocasLibBase.REALDGN rdata5 = new FocasLibBase.REALDGN();
            public FocasLibBase.REALDGN rdata6 = new FocasLibBase.REALDGN();
            public FocasLibBase.REALDGN rdata7 = new FocasLibBase.REALDGN();
            public FocasLibBase.REALDGN rdata8 = new FocasLibBase.REALDGN();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REALPRM
        {
            public int prm_val;
            public int dec_val;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REALPRMS
        {
            public FocasLibBase.REALPRM rdata1 = new FocasLibBase.REALPRM();
            public FocasLibBase.REALPRM rdata2 = new FocasLibBase.REALPRM();
            public FocasLibBase.REALPRM rdata3 = new FocasLibBase.REALPRM();
            public FocasLibBase.REALPRM rdata4 = new FocasLibBase.REALPRM();
            public FocasLibBase.REALPRM rdata5 = new FocasLibBase.REALPRM();
            public FocasLibBase.REALPRM rdata6 = new FocasLibBase.REALPRM();
            public FocasLibBase.REALPRM rdata7 = new FocasLibBase.REALPRM();
            public FocasLibBase.REALPRM rdata8 = new FocasLibBase.REALPRM();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_ALM
        {
            public short rec_type;
            public short alm_grp;
            public short alm_no;
            public sbyte axis_no;
            public sbyte dummy;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_ALM2
        {
            public short alm_grp;
            public short alm_no;
            public short axis_no;
            public short year;
            public short month;
            public short day;
            public short hour;
            public short minute;
            public short second;
            public short dummy;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_ALM2_data
        {
            public short rec_len;
            public short rec_type;
            public FocasLibBase.REC_ALM2 data = new FocasLibBase.REC_ALM2();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_ALM4
        {
            public short alm_grp;
            public short alm_no;
            public short axis_no;
            public short year;
            public short month;
            public short day;
            public short hour;
            public short minute;
            public short second;
            public short pth_no;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_ALM4_data
        {
            public short rec_len;
            public short rec_type;
            public FocasLibBase.REC_ALM4 data = new FocasLibBase.REC_ALM4();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_DATE
        {
            public short rec_type;
            public sbyte year;
            public sbyte month;
            public sbyte day;
            public sbyte pw_flag;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=2)]
            public sbyte[] dummy;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_DATE2
        {
            public short evnt_type;
            public short year;
            public short month;
            public short day;
            public short hour;
            public short minute;
            public short second;
            public short dummy;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_DATE2_data
        {
            public short rec_len;
            public short rec_type;
            public FocasLibBase.REC_DATE2 data = new FocasLibBase.REC_DATE2();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_DATE4
        {
            public short evnt_type;
            public short year;
            public short month;
            public short day;
            public short hour;
            public short minute;
            public short second;
            public short dummy;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_DATE4_data
        {
            public short rec_len;
            public short rec_type;
            public FocasLibBase.REC_DATE4 data = new FocasLibBase.REC_DATE4();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_IAL4
        {
            public short alm_grp;
            public short alm_no;
            public short axis_no;
            public short year;
            public short month;
            public short day;
            public short hour;
            public short minute;
            public short second;
            public short pth_no;
            public short sys_alm;
            public short dsp_flg;
            public short axis_num;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=10)]
            public int[] g_modal;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=10)]
            public char[] g_dp;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=10)]
            public int[] a_modal;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=10)]
            public char[] a_dp;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x20)]
            public int[] abs_pos;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x20)]
            public char[] abs_dp;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x20)]
            public int[] mcn_pos;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x20)]
            public char[] mcn_dp;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_IAL4_data
        {
            public short rec_len;
            public short rec_type;
            public FocasLibBase.REC_IAL4 data = new FocasLibBase.REC_IAL4();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_MAC4
        {
            public short mac_no;
            public short hour;
            public short minute;
            public short second;
            public short pth_no;
            public int mac_old;
            public int mac_new;
            public short old_dp;
            public short new_dp;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_MAC4_data
        {
            public short rec_len;
            public short rec_type;
            public FocasLibBase.REC_MAC4 data = new FocasLibBase.REC_MAC4();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_MAL4
        {
            public short alm_grp;
            public short alm_no;
            public short axis_no;
            public short year;
            public short month;
            public short day;
            public short hour;
            public short minute;
            public short second;
            public short pth_no;
            public short sys_alm;
            public short dsp_flg;
            public short axis_num;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x40)]
            public string alm_msg = new string(' ', 0x40);
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=10)]
            public int[] g_modal;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=10)]
            public char[] g_dp;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=10)]
            public int[] a_modal;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=10)]
            public char[] a_dp;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x20)]
            public int[] abs_pos;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x20)]
            public char[] abs_dp;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x20)]
            public int[] mcn_pos;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=0x20)]
            public char[] mcn_dp;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_MAL4_data
        {
            public short rec_len;
            public short rec_type;
            public FocasLibBase.REC_MAL4 data = new FocasLibBase.REC_MAL4();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_MDI
        {
            public short rec_type;
            public byte key_code;
            public byte pw_flag;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=4)]
            public sbyte[] dummy;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_MDI2
        {
            public byte key_code;
            public byte pw_flag;
            public short dummy;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_MDI2_data
        {
            public short rec_len;
            public short rec_type;
            public FocasLibBase.REC_MDI2 data = new FocasLibBase.REC_MDI2();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_MDI4
        {
            public char key_code;
            public char pw_flag;
            public short pth_no;
            public short ex_flag;
            public short hour;
            public short minute;
            public short second;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_MDI4_data
        {
            public short rec_len;
            public short rec_type;
            public FocasLibBase.REC_MDI4 data = new FocasLibBase.REC_MDI4();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_OFS4
        {
            public short ofs_grp;
            public short ofs_no;
            public short hour;
            public short minute;
            public short second;
            public short pth_no;
            public int ofs_old;
            public int ofs_new;
            public short old_dp;
            public short new_dp;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_OFS4_data
        {
            public short rec_len;
            public short rec_type;
            public FocasLibBase.REC_OFS4 data = new FocasLibBase.REC_OFS4();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_OPM4
        {
            public short dsp_flg;
            public short om_no;
            public short year;
            public short month;
            public short day;
            public short hour;
            public short minute;
            public short second;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x100)]
            public string ope_msg = new string(' ', 0x100);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_OPM4_data
        {
            public short rec_len;
            public short rec_type;
            public FocasLibBase.REC_OPM4 data = new FocasLibBase.REC_OPM4();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_PRM4
        {
            public short prm_grp;
            public short prm_num;
            public short hour;
            public short minute;
            public short second;
            public short prm_len;
            public int prm_no;
            public int prm_old;
            public int prm_new;
            public short old_dp;
            public short new_dp;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_PRM4_data
        {
            public short rec_len;
            public short rec_type;
            public FocasLibBase.REC_PRM4 data = new FocasLibBase.REC_PRM4();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_SGN
        {
            public short rec_type;
            public sbyte sig_name;
            public byte sig_old;
            public byte sig_new;
            public sbyte dummy;
            public short sig_no;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_SGN2
        {
            public short sig_name;
            public short sig_no;
            public byte sig_old;
            public byte sig_new;
            public short dummy;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_SGN2_data
        {
            public short rec_len;
            public short rec_type;
            public FocasLibBase.REC_SGN2 data = new FocasLibBase.REC_SGN2();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_SGN4
        {
            public short sig_name;
            public short sig_no;
            public char sig_old;
            public char sig_new;
            public short pmc_no;
            public short hour;
            public short minute;
            public short second;
            public short dummy;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_SGN4_data
        {
            public short rec_len;
            public short rec_type;
            public FocasLibBase.REC_SGN4 data = new FocasLibBase.REC_SGN4();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_TIME
        {
            public short rec_type;
            public sbyte hour;
            public sbyte minute;
            public sbyte second;
            public sbyte pw_flag;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=2)]
            public sbyte[] dummy;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_WOF4
        {
            public short ofs_grp;
            public short ofs_no;
            public short hour;
            public short minute;
            public short second;
            public short pth_no;
            public short axis_no;
            public short dummy;
            public int ofs_old;
            public int ofs_new;
            public short old_dp;
            public short new_dp;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class REC_WOF4_data
        {
            public short rec_len;
            public short rec_type;
            public FocasLibBase.REC_WOF4 data = new FocasLibBase.REC_WOF4();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class SPEEDELM
        {
            public int data;
            public short dec;
            public short unit;
            public short disp;
            public byte name;
            public byte suff;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class SRAMADDR
        {
            public short type;
            public int size;
            public int offset;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class T_OFS_A
        {
            public short tip;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=4)]
            public int[] data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class T_OFS_A_data
        {
            public FocasLibBase.T_OFS_A data1 = new FocasLibBase.T_OFS_A();
            public FocasLibBase.T_OFS_A data2 = new FocasLibBase.T_OFS_A();
            public FocasLibBase.T_OFS_A data3 = new FocasLibBase.T_OFS_A();
            public FocasLibBase.T_OFS_A data4 = new FocasLibBase.T_OFS_A();
            public FocasLibBase.T_OFS_A data5 = new FocasLibBase.T_OFS_A();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class T_OFS_B
        {
            public short tip;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)]
            public int[] data;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class T_OFS_B_data
        {
            public FocasLibBase.T_OFS_B data1 = new FocasLibBase.T_OFS_B();
            public FocasLibBase.T_OFS_B data2 = new FocasLibBase.T_OFS_B();
            public FocasLibBase.T_OFS_B data3 = new FocasLibBase.T_OFS_B();
            public FocasLibBase.T_OFS_B data4 = new FocasLibBase.T_OFS_B();
            public FocasLibBase.T_OFS_B data5 = new FocasLibBase.T_OFS_B();
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public class TCPPRM
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x10)]
            public string OwnIPAddress = new string(' ', 0x10);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x10)]
            public string SubNetMask = new string(' ', 0x10);
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=0x10)]
            public string RouterIPAddress = new string(' ', 0x10);
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public struct TIMER_DATE
        {
            public short year;
            public short month;
            public short date;
        }

        [StructLayout(LayoutKind.Sequential, Pack=4)]
        public struct TIMER_TIME
        {
            public short hour;
            public short minute;
            public short second;
        }
                

        [StructLayout(LayoutKind.Explicit, Pack = 4)]
        public class ODBOPHIS4_SATYA
        {
            [FieldOffset(0)]
            public short rec_len;
            [FieldOffset(2)]
            public short rec_type;

            [FieldOffset(4)]
            public FocasLibBase.REC_MDI4 REC_MDI4 = new FocasLibBase.REC_MDI4();
            [FieldOffset(4)]
            public FocasLibBase.REC_SGN4 REC_SGN4 = new FocasLibBase.REC_SGN4();
            [FieldOffset(4)]
            public FocasLibBase.REC_ALM4 REC_ALM4 = new FocasLibBase.REC_ALM4();
            [FieldOffset(4)]
            public FocasLibBase.REC_DATE4 REC_DATE4 = new FocasLibBase.REC_DATE4();
            [FieldOffset(4)]
            public FocasLibBase.REC_IAL4 REC_IAL4 = new FocasLibBase.REC_IAL4();
            [FieldOffset(4)]
            public FocasLibBase.REC_MAL4 REC_MAL4 = new FocasLibBase.REC_MAL4();
            [FieldOffset(4)]
            public FocasLibBase.REC_OPM4 REC_OPM4 = new FocasLibBase.REC_OPM4();
            [FieldOffset(4)]
            public FocasLibBase.REC_OFS4 REC_OFS4 = new FocasLibBase.REC_OFS4();
            [FieldOffset(4)]
            public FocasLibBase.REC_PRM4 REC_PRM4 = new FocasLibBase.REC_PRM4();
            [FieldOffset(4)]
            public FocasLibBase.REC_WOF4 REC_WOF4 = new FocasLibBase.REC_WOF4();
            [FieldOffset(4)]
            public FocasLibBase.REC_MAC4 REC_MAC4 = new FocasLibBase.REC_MAC4();
        }


    }
}


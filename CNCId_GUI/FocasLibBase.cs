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

}


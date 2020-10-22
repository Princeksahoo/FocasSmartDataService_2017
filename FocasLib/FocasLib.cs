using System;
using System.Runtime.InteropServices;

namespace FocasLibrary
{
    public class FocasLib : FocasLibBase
    {
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_absolute(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBAXIS c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_absolute2(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBAXIS c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_accdecdly(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBAXIS c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_actf(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBACT a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_acts(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBACT a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_acts2(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBACT2 b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_adcnv(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBAD c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_alarm(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBALM a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_alarm2(ushort FlibHndl, out int a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_allclibhndl(out ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_allclibhndl2(int node, out ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_allclibhndl3([In, MarshalAs(UnmanagedType.AsAny)] object ip, ushort port, int timeout, out ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_allclibhndl4([In, MarshalAs(UnmanagedType.AsAny)] object ip, ushort port, int timeout, uint id, out ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_allowance(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBAXIS c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_allowcnd(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBCAXIS c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_async_busy_state(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_buff(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBBUF a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_canmovrlap(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_cdnc(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, ushort b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_cdownload(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, short b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_cexedirectory(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, ref ushort b, ushort c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.CFILEINFO d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_cexesraminfo(ushort FlibHndl, out short a, out int b, out int c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_cexesramsize(ushort FlibHndl, out int a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_chgprotbit(ushort FlibHndl, short a, ref byte b, int c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_chkrmtdgn(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_clearomhis(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_clearophis(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_closecexefile(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_clr3dplsmov(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_clr5axpls(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_clralm(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_clrcntinfo(ushort FlibHndl, short a, short b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_clrdgdat(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_clrgrphcmd(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_clrmsgbuff(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_condense(ushort FlibHndl, short a, int b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_copyprog(ushort FlibHndl, int a, int b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_cupload(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBUP a, ref ushort b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_cverify(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, short b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_delall(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_delete(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_delmagazine(ushort FlibHndl, ref short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTLMAG2 b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_delprogline(ushort FlibHndl, int a, uint b, uint c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_delrange(ushort FlibHndl, int a, int b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_deltlifedt(ushort FlibHndl, short a, short b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_deltlifegrp(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_deltool(ushort FlibHndl, short a, ref short b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_diagnosr(ushort FlibHndl, ref short a, short b, ref short c, ref short d, [Out, MarshalAs(UnmanagedType.AsAny)] object e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_diagnoss(ushort FlibHndl, short a, short b, short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBDGN_1 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_diagnoss(ushort FlibHndl, short a, short b, short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBDGN_2 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_diagnoss(ushort FlibHndl, short a, short b, short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBDGN_3 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_diagnoss(ushort FlibHndl, short a, short b, short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBDGN_4 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dispoptmsg(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_distance(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBAXIS c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dnc(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, ushort b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dnc2(ushort FlibHndl, ref int a, [In, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dncend(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dncend2(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dncstart(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dncstart2(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_download(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, short b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_download3(ushort FlibHndl, ref int a, [In, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_download4(ushort FlibHndl, ref int a, [In, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dtsvchkdsk(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dtsvcnclupdn(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dtsvdelete(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dtsvdownload(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dtsvftpget(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, [In, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dtsvftpput(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, [In, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dtsvftpstat(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dtsvgetdncpg(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dtsvgetmode(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dtsvhdformat(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dtsvmntinfo(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBDSMNT a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dtsvrdcram(ushort FlibHndl, int a, ref int b, [Out, MarshalAs(UnmanagedType.AsAny)] object c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dtsvrderrmsg(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dtsvrdfile(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, [In, MarshalAs(UnmanagedType.AsAny)] object b, short c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dtsvrdpgdir(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBDSDIR c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dtsvrdset(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBDSSET a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dtsvsavecram(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dtsvsetdncpg(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dtsvsetmode(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dtsvupdnstat(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dtsvupload(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dtsvwrfile(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, [In, MarshalAs(UnmanagedType.AsAny)] object b, short c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dtsvwrset(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBDSSET a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dwnend(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dwnend2(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dwnend3(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dwnend4(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dwnstart(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dwnstart3(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dwnstart3_f(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.AsAny)] object b, [In, MarshalAs(UnmanagedType.AsAny)] object c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_dwnstart4(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_end_async_wrparam(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_exaxisname(ushort FlibHndl, short a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBEXAXISNAME c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_exeprgname(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBEXEPRG a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_freelibhndl(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_fromdelete(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.AsAny)] object b, int c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_fromget(ushort FlibHndl, out short a, [Out, MarshalAs(UnmanagedType.AsAny)] object b, ref int c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_fromgetend(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_fromgetstart(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_fromldend(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_fromldstart(ushort FlibHndl, short a, int b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_fromload(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, ref int b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_fromput(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, ref int b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_fromputend(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_fromputstart(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_fromremove(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_fromsave(ushort FlibHndl, out short a, [Out, MarshalAs(UnmanagedType.AsAny)] object b, ref int c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_fromsvend(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_fromsvstart(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.AsAny)] object b, int c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_ftosjis(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, [Out, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_getcncmodel(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_getcrntscrn(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_getdspmode(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_getdtailerr(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBERR a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_getfigure(ushort FlibHndl, short a, out short b, [Out, MarshalAs(UnmanagedType.AsAny)] object c, [Out, MarshalAs(UnmanagedType.AsAny)] object d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_getfocas1opt(ushort FlibHndl, short a, out int b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_getfrominfo(ushort FlibHndl, short a, out short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBFINFORM c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_getlibopt(ushort FlibHndl, int a, [Out, MarshalAs(UnmanagedType.AsAny)] object b, ref int c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_getlockstat(ushort FlibHndl, short a, out byte b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_getmactype(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_getpath(ushort FlibHndl, out short a, out short b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_getpmactype(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_getsraminfo(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSINFO a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_gettimer(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTIMER a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_hpccactfine(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_hpccatset(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_hpccattune(ushort FlibHndl, short a, out short b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_hpccselfine(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_instlifedt(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IDBITD a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_machine(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBAXIS c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_mergeprog(ushort FlibHndl, short a, int b, uint c, int d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_modal(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBMDL_1 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_modal(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBMDL_2 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_modal(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBMDL_3 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_modal(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBMDL_4 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_modal(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBMDL_5 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_newprog(ushort FlibHndl, int a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_opencexefile(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, short b, short c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_optmsgans(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_pdf_add(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_pdf_cond(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_pdf_copy(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, [In, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_pdf_del(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_pdf_delline(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, uint b, uint c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_pdf_move(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, [In, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_pdf_rdactpt(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.AsAny)] object a, out int b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_pdf_rdmain(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_pdf_rename(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, [In, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_pdf_searchresult(ushort FlibHndl, out uint a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_pdf_searchword(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, uint b, uint c, uint d, uint e, [In, MarshalAs(UnmanagedType.AsAny)] object f);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_pdf_slctmain(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_pdf_wractpt(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, short b, ref int c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_progdigit(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_prstwkcd(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IDBWRA b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rd1length(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBTLIFE4 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rd1radius(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBTLIFE4 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rd1tlifedat2(ushort FlibHndl, short a, int b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTD2 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rd1tlifedata(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTD c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rd2length(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBTLIFE4 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rd2radius(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBTLIFE4 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rd2tlifedata(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTD c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rd3dcdcnv(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODB3DCD a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rd3dmovrlap(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODB3DHDL a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rd3dofschg(ushort FlibHndl, ref int a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rd3dpulse(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODB3DPLS a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rd3dtofs(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODB3DTO a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rd3dtooltip(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODB3DHDL a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rd5axmandt(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODB5AXMAN a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rd5axovrlap(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBAXIS c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdactfixofs(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBZOFS b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdactpt(ushort FlibHndl, out int a, out int b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdacttlzone(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBACTTLZN a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdalmhisno(ushort FlibHndl, out ushort a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdalmhistry(ushort FlibHndl, ushort a, ushort b, ushort c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBAHIS d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdalmhistry_w(ushort FlibHndl, ushort a, ushort b, ushort c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBAHIS d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdalmhistry2(ushort FlibHndl, ushort a, ushort b, ushort c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBAHIS2 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdalmhistry3(ushort FlibHndl, ushort a, ushort b, ushort c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBAHIS3 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdalmhistry5(ushort FlibHndl, ushort a, ushort b, ushort c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBAHIS5 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdalminfo(ushort FlibHndl, short a, short b, short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ALMINFO_1 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdalminfo(ushort FlibHndl, short a, short b, short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ALMINFO_2 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdalmmsg(ushort FlibHndl, short a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBALMMSG c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdalmmsg2(ushort FlibHndl, short a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBALMMSG2 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdaxisdata(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.AsAny)] object b, short c, ref short d, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBAXDT e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdaxisname(ushort FlibHndl, ref short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBAXISNAME b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdbaxis(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBBAXIS a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdblkcount(ushort FlibHndl, out int a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdbrstrinfo(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBBRS a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdbtofsinfo(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBBTLINF a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdbtofsr(ushort FlibHndl, short a, short b, short c, short d, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBBTO e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdcdrotate(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBROT a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdcexefile(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.AsAny)] object a, ref uint b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdcexesram(ushort FlibHndl, int a, [In, MarshalAs(UnmanagedType.AsAny)] object b, ref int c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdcomlogmsg(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdcommand(ushort FlibHndl, short a, short b, ref short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBCMD d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdcomopemsg(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdcomparam(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBCPRM a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdcoordnum(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdcount(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBTLIFE3 b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdctrldi(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSPDI a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdctrldo(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSPDO a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdcurrent(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rddgdat(ushort FlibHndl, ref short a, [Out, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rddiag_ext(ushort FlibHndl, [In, Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPRMNO a, short b, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPRM c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rddiaginfo(ushort FlibHndl, short a, ushort b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBDIAGIF c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rddiagnum(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBDIAGNUM a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rddischarge(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBDISCHRG a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rddischrgalm(ushort FlibHndl, int a, ref int b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBDISCHRGALM c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rddncdgndt(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBDNCDGN a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rddncfname(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rddynamic(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBDY_1 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rddynamic(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBDY_2 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rddynamic2(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBDY2_1 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rddynamic2(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBDY2_2 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdetherinfo(ushort FlibHndl, out short a, out short b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdexchgtgrp(ushort FlibHndl, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBEXGP c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdexecmcode(ushort FlibHndl, short a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBEXEM c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdexecprog(ushort FlibHndl, ref ushort a, out short b, [Out, MarshalAs(UnmanagedType.AsAny)] object c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdexecpt(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.PRGPNT a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.PRGPNT b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdfbusmem(ushort FlibHndl, short a, short b, int c, int d, [Out, MarshalAs(UnmanagedType.AsAny)] object e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdfixcycle(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBFIX a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdfixofs(ushort FlibHndl, short a, short b, short c, short d, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBZOR e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdfrominfo(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBFINFO b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdgcode(ushort FlibHndl, short a, short b, ref short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBGCD d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdgrphcanflg(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdgrphcmd(ushort FlibHndl, ref short a, [Out, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdgrpid(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBTLIFE1 b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdgrpid2(ushort FlibHndl, int a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBTLIFE5 b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdgrpinfo(ushort FlibHndl, short a, short b, short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTGI d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdgrpinfo2(ushort FlibHndl, short a, short b, short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTGI2 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdgrpinfo3(ushort FlibHndl, short a, short b, short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTGI3 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdgrpinfo4(ushort FlibHndl, short a, short b, short c, out short d, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTGI4 e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdhissgnl(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBSIG a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdhissgnl2(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBSIG2 a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdhissgnl3(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBSIG3 a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdhndintrpt(ushort FlibHndl, short a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBHND c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdhpccset(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBHPST a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdhpcctuac(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBHPAC a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdhpcctupr(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBHPPR a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdhsparam(ushort FlibHndl, int a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.HSPINFO b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.HSPDATA_1 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdhsparam(ushort FlibHndl, int a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.HSPINFO b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.HSPDATA_2 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdhsparam(ushort FlibHndl, int a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.HSPINFO b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.HSPDATA_3 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdhsprminfo(ushort FlibHndl, int a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.HSPINFO_data b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdintchk(ushort FlibHndl, short a, short b, short c, short d, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBINT e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdintinfo(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdlactnum(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBLACTN a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdlagslt(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBLAGSL a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdlagst(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBLAGST a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdlcmddat(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBLCMDT a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdlcmmt(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBLCMMT a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdldsplc(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdledgprc(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBLEGPR a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdlenofs(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBLOFS a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdlerrz(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdlife(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBTLIFE3 b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdloopgain(ushort FlibHndl, out int a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdlprcprc(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBLPCPR a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdlpwrcpst(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdlpwrctrl(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdlpwrdat(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBLOPDT a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdlpwrdty(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBLPWDT a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdlpwrslt(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdmacro(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBM c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdmacroinfo(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBMVINF a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdmacror(ushort FlibHndl, short a, short b, short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBMR d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdmacror2(ushort FlibHndl, int a, ref int b, [Out, MarshalAs(UnmanagedType.AsAny)] object c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdmagazine(ushort FlibHndl, ref short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTLMAG b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdmaxgrp(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBLFNO a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdmaxtool(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBLFNO a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdmdipntr(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBMDIP a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdmdiprgstat(ushort FlibHndl, out ushort a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdmdiprog(ushort FlibHndl, ref short a, [Out, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdmdlconfig(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBMDLC a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdmdlconfig2(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdmenuswitch(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdmgrpdata(ushort FlibHndl, short a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBMGRP c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdmirimage(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBMIR a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdmngtime(ushort FlibHndl, int a, ref int b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBMNGTIME c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdmovrlap(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBOVL c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdmsptype(ushort FlibHndl, short a, short b, short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBMSTP d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdmtapdata(ushort FlibHndl, short a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBMTAP c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdmultipieceno(ushort FlibHndl, out int a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdmultitldt(ushort FlibHndl, short a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBMLTTL c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdngrp(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBTLIFE2 a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdnodeinfo(int a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBNODE b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdnodenum(out int a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdnspdl(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdntool(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBTLIFE3 b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdomhisinfo(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBOMIF a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdomhisno(ushort FlibHndl, out ushort a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdomhistry(ushort FlibHndl, ushort a, ref ushort b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBOMHIS c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdomhistry2(ushort FlibHndl, ushort a, ushort b, ushort c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBOMHIS2 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdophisno(ushort FlibHndl, out ushort a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdophistry(ushort FlibHndl, ushort a, ushort b, ushort c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBHIS d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdophistry2(ushort FlibHndl, ushort a, ref ushort b, ref ushort c, [Out, MarshalAs(UnmanagedType.AsAny)] object d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdophistry4(ushort FlibHndl, ushort a, ref ushort b, ref ushort c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBOPHIS4_1 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdophistry4(ushort FlibHndl, ushort a, ref ushort b, ref ushort c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBOPHIS4_10 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdophistry4(ushort FlibHndl, ushort a, ref ushort b, ref ushort c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBOPHIS4_11 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdophistry4(ushort FlibHndl, ushort a, ref ushort b, ref ushort c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBOPHIS4_2 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdophistry4(ushort FlibHndl, ushort a, ref ushort b, ref ushort c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBOPHIS4_3 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdophistry4(ushort FlibHndl, ushort a, ref ushort b, ref ushort c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBOPHIS4_4 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdophistry4(ushort FlibHndl, ushort a, ref ushort b, ref ushort c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBOPHIS4_5 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdophistry4(ushort FlibHndl, ushort a, ref ushort b, ref ushort c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBOPHIS4_6 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdophistry4(ushort FlibHndl, ushort a, ref ushort b, ref ushort c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBOPHIS4_7 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdophistry4(ushort FlibHndl, ushort a, ref ushort b, ref ushort c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBOPHIS4_8 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdophistry4(ushort FlibHndl, ushort a, ref ushort b, ref ushort c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBOPHIS4_9 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdopmode(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdopmsg(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.OPMSG c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdopmsg2(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.OPMSG2 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdopmsg3(ushort FlibHndl, short a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.OPMSG3 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdopnlgnrl(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBGNRL b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdopnlgsname(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBRDNA b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdopnlsgnl(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBSGNL b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdparainfo(ushort FlibHndl, short a, ushort b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBPARAIF c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdparam(ushort FlibHndl, short a, short b, short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPSD_1 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdparam(ushort FlibHndl, short a, short b, short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPSD_2 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdparam(ushort FlibHndl, short a, short b, short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPSD_3 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdparam(ushort FlibHndl, short a, short b, short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPSD_4 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdparam_ext(ushort FlibHndl, [In, Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPRMNO a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPRM c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdparam3(ushort FlibHndl, short a, short b, short c, short d, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPSD_1 e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdparam3(ushort FlibHndl, short a, short b, short c, short d, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPSD_2 e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdparam3(ushort FlibHndl, short a, short b, short c, short d, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPSD_3 e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdparam3(ushort FlibHndl, short a, short b, short c, short d, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPSD_4 e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdparanum(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBPARANUM a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdparar(ushort FlibHndl, ref short a, short b, ref short c, ref short d, [Out, MarshalAs(UnmanagedType.AsAny)] object e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdpdf_alldir(ushort FlibHndl, ref short a, [In, MarshalAs(UnmanagedType.AsAny)] object b, [Out, MarshalAs(UnmanagedType.AsAny)] object c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdpdf_curdir(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdpdf_drive(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdpdf_inf(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, short b, [Out, MarshalAs(UnmanagedType.AsAny)] object c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdpdf_line(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, uint b, [Out, MarshalAs(UnmanagedType.AsAny)] object c, ref uint d, ref uint e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdpdf_subdir(ushort FlibHndl, ref short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IDBPDFSDIR b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBPDFSDIR c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdpdf_subdirn(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBPDFNFIL b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdpglockstat(ushort FlibHndl, out int a, out int b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdpitchinfo(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdpitchr(ushort FlibHndl, short a, short b, short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPI d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdpm_cncitem(ushort FlibHndl, short a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBITEM c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdpm_item(ushort FlibHndl, short a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPMAINTE c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdpm_mcnitem(ushort FlibHndl, short a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBITEM c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdpmacro(ushort FlibHndl, int a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBPM b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdpmacroinfo(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBPMINF a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdpmacror(ushort FlibHndl, int a, int b, ushort c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPR d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdpmacror2(ushort FlibHndl, uint a, ref uint b, ushort c, [Out, MarshalAs(UnmanagedType.AsAny)] object d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdposerrs(ushort FlibHndl, out int a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdposerrs2(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBPSER a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdposerrz(ushort FlibHndl, out int a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdposition(ushort FlibHndl, short a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBPOS c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdposofs(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBPOFS a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdprgdirtime(ushort FlibHndl, ref int a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.PRGDIRTM c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdprgnum(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBPRO a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdproctime(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBPTIME a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdprogdir(ushort FlibHndl, short a, short b, short c, ushort d, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.PRGDIR e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdprogdir2(ushort FlibHndl, short a, ref short b, ref short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.PRGDIR2 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdprogdir3(ushort FlibHndl, short a, ref int b, ref short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.PRGDIR3 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdprogdir4(ushort FlibHndl, short a, int b, ref short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.PRGDIR4 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdproginfo(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBNC_1 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdproginfo(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBNC_2 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdprogline(ushort FlibHndl, int a, uint b, [Out, MarshalAs(UnmanagedType.AsAny)] object c, ref uint d, ref uint e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdprogline2(ushort FlibHndl, int a, uint b, [Out, MarshalAs(UnmanagedType.AsAny)] object c, ref uint d, ref uint e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdprstrinfo(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBPRS a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdpscdedge(ushort FlibHndl, short a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBEDGE c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdpscdpirc(ushort FlibHndl, short a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPIRC c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdpscdproc(ushort FlibHndl, short a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPSCD c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdpscdslop(ushort FlibHndl, short a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBSLOP c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdpwofsthis(ushort FlibHndl, int a, ref int b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBPWOFST c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdradofs(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBROFS a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdrcvmsg(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdrcvstat(ushort FlibHndl, out ushort a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdrepeatval(ushort FlibHndl, out int a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdrmtdgn(ushort FlibHndl, out int a, [Out, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdrmtwavedt(ushort FlibHndl, short a, int b, ref int c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBRMTDT d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdrmtwaveprm(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBRMTPRM a, short b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdrstrmcode(ushort FlibHndl, short a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBRSTRM c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdsafetyzone(ushort FlibHndl, short a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBSAFE c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdsavsigadr(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBSIGAD a, short b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdsavsigdata(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.AsAny)] object c, ref short d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdscaling(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSCL a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdseqnum(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSEQ a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdset(ushort FlibHndl, short a, short b, short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPSD_1 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdset(ushort FlibHndl, short a, short b, short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPSD_2 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdset(ushort FlibHndl, short a, short b, short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPSD_3 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdset(ushort FlibHndl, short a, short b, short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPSD_4 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdsetinfo(ushort FlibHndl, short a, ushort b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSETIF c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdsetnum(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSETNUM a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdsetr(ushort FlibHndl, ref short a, short b, ref short c, ref short d, [Out, MarshalAs(UnmanagedType.AsAny)] object e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdsetzone(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdsndmsg(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdspcss(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBCSS a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdspdlalm(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdspdlname(ushort FlibHndl, ref short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSPDLNAME b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdspeed(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSPEED b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdspgear(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSPN b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdspload(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSPN b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdspmaxrpm(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSPN b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdspmeter(ushort FlibHndl, short a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSPLOAD c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdsramaddr(ushort FlibHndl, out short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.SRAMADDR b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdsraminfo(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSINFO a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdsramnum(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdsrvspeed(ushort FlibHndl, out int a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdsvmeter(ushort FlibHndl, ref short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSVLOAD b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdsynerrrg(ushort FlibHndl, out int a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdsynerrsy(ushort FlibHndl, out int a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdsyshard(ushort FlibHndl, short a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSYSH c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdsyssoft(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSYSS a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdsyssoft2(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSYSS2 a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdsyssoft3(ushort FlibHndl, short a, ref short b, out short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSYSS3 d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdtimer(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTIME b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdtlctldata(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTLCTL a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdtlgrp(ushort FlibHndl, int a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTLGRP c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdtlinfo(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBTLINFO a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdtltool(ushort FlibHndl, int a, int b, ref short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTLTOOL d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdtlusegrp(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBUSEGRP a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdtofs(ushort FlibHndl, short a, short b, short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBTOFS d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdtofsinfo(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBTLINF a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdtofsinfo2(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBTLINF2 a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdtofsr(ushort FlibHndl, short a, short b, short c, short d, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTO_1_1 e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdtofsr(ushort FlibHndl, short a, short b, short c, short d, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTO_1_2 e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdtofsr(ushort FlibHndl, short a, short b, short c, short d, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTO_1_3 e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdtofsr(ushort FlibHndl, short a, short b, short c, short d, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTO_2 e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdtofsr(ushort FlibHndl, short a, short b, short c, short d, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTO_3 e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdtool(ushort FlibHndl, short a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTLMNG c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdtool_f2(ushort FlibHndl, short a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTLMNG_F2 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdtooldata(ushort FlibHndl, short a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTLDT c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdtoolgrp(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBTG c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdtoolinfo(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBPTLINF a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdtoolrng(ushort FlibHndl, short a, short b, short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTR d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdtoolzone(ushort FlibHndl, short a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTLZN c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdunsolicmsg(short a, [In, Out] FocasLibBase.IDBUNSOLICMSG b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdunsolicprm(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBUNSOLIC b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdusegrpid(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBUSEGR a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdusetlno(ushort FlibHndl, short a, short b, short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBTLUSE d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rduvactpt(ushort FlibHndl, out int a, out int b, out int c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdwavedata(ushort FlibHndl, short a, short b, int c, ref int d, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBWVDT e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdwaveprm(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBWAVE a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdwaveprm2(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBWVPRM a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdwkcdsfms(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBWCSF c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdwkcdshft(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBWCSF c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdzofs(ushort FlibHndl, short a, short b, short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBZOFS d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdzofsinfo(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rdzofsr(ushort FlibHndl, short a, short b, short c, short d, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBZOR e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_regmagazine(ushort FlibHndl, ref short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTLMAG b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_regtool(ushort FlibHndl, short a, ref short b, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTLMNG c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_regtool_f2(ushort FlibHndl, short a, ref short b, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTLMNG_F2 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_relative(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBAXIS c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_relative2(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBAXIS c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_renameprog(ushort FlibHndl, int a, int b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_reset(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_reset2(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_resetconnect(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_resetpglock(ushort FlibHndl, int a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rewind(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rmtwavestart(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rmtwavestat(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rmtwavestop(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rstrseqsrch(ushort FlibHndl, int a, int b, short c, short d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_rstrseqsrch2(ushort FlibHndl, int a, int b, short c, short d, int e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_sdcancelsmpl(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_sdclrchnl(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_sdendsmpl(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_sdread1shot(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_sdreadsmpl(ushort FlibHndl, out short a, int b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSD c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_sdsetchnl(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IDBCHAN b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_sdstartsmpl(ushort FlibHndl, short a, int b, [Out, MarshalAs(UnmanagedType.AsAny)] object c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_search(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_search2(ushort FlibHndl, int a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_searchresult(ushort FlibHndl, out uint a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_searchword(ushort FlibHndl, int a, uint b, short c, short d, uint e, [In, MarshalAs(UnmanagedType.AsAny)] object f);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_sendmessage(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_seqsrch(ushort FlibHndl, int a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_seqsrch2(ushort FlibHndl, int a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_setdefnode(int a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_setlibopt(ushort FlibHndl, int a, [In, MarshalAs(UnmanagedType.AsAny)] object b, int c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_setmactype(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_setpath(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_setpglock(ushort FlibHndl, int a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_setpmactype(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_setthrdngpos(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_settimeout(ushort FlibHndl, int a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_settimer(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTIMER a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_setvrtclpos(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_sfbcancelsmpl(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_sfbclrchnl(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_sfbendsmpl(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_sfbreadsmpl(ushort FlibHndl, out short a, int b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSD c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_sfbsetchnl(ushort FlibHndl, short a, int b, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IDBSFBCHAN c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_sfbstartsmpl(ushort FlibHndl, short a, int b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_skip(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBAXIS c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_slctscrn(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_slide(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBAXIS c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_srambackup(ushort FlibHndl, out short a, [Out, MarshalAs(UnmanagedType.AsAny)] object b, ref int c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_srambkend(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_srambkstart(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, int b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_sramget(ushort FlibHndl, out short a, [Out, MarshalAs(UnmanagedType.AsAny)] object b, ref int c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_sramget2(ushort FlibHndl, out short a, [Out, MarshalAs(UnmanagedType.AsAny)] object b, ref int c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_sramgetend(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_sramgetend2(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_sramgetstart(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_sramgetstart2(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_srcsfreechnl(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_srcsrddrvcp(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_srcsrdexstat(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSRCSST a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_srcsrdidinfo(ushort FlibHndl, int a, short b, short c, [Out, MarshalAs(UnmanagedType.AsAny)] FocasLibBase.IODBIDINF d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_srcsrdlayout(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSRCSLYT a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_srcsrdopdata(ushort FlibHndl, int a, ref int b, [Out, MarshalAs(UnmanagedType.AsAny)] object c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_srcsrsvchnl(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_srcsstartrd(ushort FlibHndl, int a, short b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_srcsstartwrt(ushort FlibHndl, int a, short b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_srcsstopexec(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_srcswridinfo(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBIDINF a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_srcswropdata(ushort FlibHndl, int a, int b, [In, MarshalAs(UnmanagedType.AsAny)] object c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_srvdelay(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBAXIS c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_start(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_start_async_wrparam(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPRM a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_startdrawpos(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_startdyngrph(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_startgetdgdat(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_startnccmd(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_startnccmd2(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_startomhis(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_startophis(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_startrmtdgn(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_statinfo(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBST a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_stopdrawpos(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_stopdyngrph(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_stopgetdgdat(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_stopnccmd(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_stopomhis(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_stopophis(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_stoprmtdgn(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_svdtendrd(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_svdtendwr(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_svdtrddata(ushort FlibHndl, out short a, ref int b, [Out, MarshalAs(UnmanagedType.AsAny)] object c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_svdtstartrd(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_svdtstartwr(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_svdtstopexec(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_svdtwrdata(ushort FlibHndl, out short a, ref int b, [In, MarshalAs(UnmanagedType.AsAny)] object c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_sysconfig(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSYSC a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_sysinfo(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSYS a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_sysinfo_ex(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSYSEX a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_t1info(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBTLIFE4 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_t2info(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBTLIFE4 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_tofs_rnge(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBDATRNG c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_toolnum(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBTLIFE4 c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_unsolicstart(ushort FlibHndl, short a, int hWnd, uint c, short d, out short e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_unsolicstop(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_upend(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_upend3(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_upend4(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_upload(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBUP a, ref ushort b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_upload3(ushort FlibHndl, ref int a, [Out, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_upload4(ushort FlibHndl, ref int a, [Out, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_upstart(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_upstart3(ushort FlibHndl, short a, int b, int c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_upstart3_f(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.AsAny)] object b, [In, MarshalAs(UnmanagedType.AsAny)] object c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_upstart4(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_verify(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, short b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_verify4(ushort FlibHndl, ref int a, [In, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_vrfend(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_vrfend4(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_vrfstart(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_vrfstart4(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wavestart(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wavestat(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wavestop(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wksft_rnge(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBDATRNG b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_workzero(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBZOFS b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wr1tlifedat2(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTD2 a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wr1tlifedata(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTD a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wr2tlifedata(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTD a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wractpt(ushort FlibHndl, int a, short b, ref int c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrbtofsr(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBBTO b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrcexefile(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, ref uint b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrcexesram(ushort FlibHndl, int a, [In, MarshalAs(UnmanagedType.AsAny)] object b, ref int c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrcommstatus(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrcomparam(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBCPRM a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrcountr(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IDBWRC b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrdgdatptr(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrdncfname(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrfbusmem(ushort FlibHndl, short a, short b, int c, int d, [In, MarshalAs(UnmanagedType.AsAny)] object e);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrfixofs(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBZOR b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrgrphcmdptr(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrgrpinfo(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTGI b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrgrpinfo2(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTGI2 b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrgrpinfo3(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTGI3 b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrhissgnl(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBSIG a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrhissgnl2(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBSIG2 a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrhissgnl3(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBSIG3 a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrhpccset(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBHPST a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrhpcctuac(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBHPAC a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrhpcctupr(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBHPPR a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrintchk(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBINT b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrkeyhistry(ushort FlibHndl, byte a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrlagslt(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBLAGSL a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrlagst(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBLAGST a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrldsplc(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrledgprc(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBLEGPR a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrlprcprc(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBLPCPR a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrlpwrcpst(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrlpwrctrl(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrlpwrdty(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBLPWDT a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrlpwrslt(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrmacro(ushort FlibHndl, short a, short b, int c, short d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrmacror(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBMR b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrmacror2(ushort FlibHndl, int a, ref int b, [In, MarshalAs(UnmanagedType.AsAny)] object c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrmagazine(ushort FlibHndl, short a, short b, short c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrmdipntr(ushort FlibHndl, int a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrmdiprog(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrmenuswitch(ushort FlibHndl, short a, short b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrmgrpdata(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IDBMGRP a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrmngtime(ushort FlibHndl, int a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBMNGTIME b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrmsptype(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBMSTP d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrmtapdata(ushort FlibHndl, short a, ref short b, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBMTAP c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrmultitldt(ushort FlibHndl, short a, ref short b, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBMLTTL c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wropnlgnrl(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBGNRL a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wropnlgsname(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBRDNA a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wropnlsgnl(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBSGNL a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrparam(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPSD_1 b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrparam(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPSD_2 b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrparam(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPSD_3 b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrparam(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPSD_4 b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrparas(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrpdf_attr(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IDBPDFTDIR b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrpdf_curdir(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrpdf_line(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, uint b, [In, MarshalAs(UnmanagedType.AsAny)] object c, uint d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrpitchr(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPI b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrpm_item(ushort FlibHndl, short a, short b, short c, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPMAINTE d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrpm_mcnitem(ushort FlibHndl, short a, short b, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBITEM c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrpmacro(ushort FlibHndl, int a, int b, short c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrpmacror(ushort FlibHndl, ushort a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPR b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrpmacror2(ushort FlibHndl, uint a, ref uint b, ushort c, [In, MarshalAs(UnmanagedType.AsAny)] object d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrprogline(ushort FlibHndl, int a, uint b, [In, MarshalAs(UnmanagedType.AsAny)] object c, uint d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrpscdedge(ushort FlibHndl, short a, ref short b, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBEDGE c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrpscdpirc(ushort FlibHndl, short a, ref short b, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPIRC c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrpscdproc(ushort FlibHndl, short a, ref short b, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPSCD c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrpscdslop(ushort FlibHndl, short a, ref short b, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBSLOP c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrrelpos(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IDBWRR b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrrmtdgn(ushort FlibHndl, ref int a, [In, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrrmtwaveprm(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBRMTPRM a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrsafetyzone(ushort FlibHndl, short a, ref short b, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBSAFE c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrsavsigadr(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBSIGAD a, out short b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrset(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPSD_1 b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrset(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPSD_2 b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrset(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPSD_3 b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrset(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPSD_4 b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrsets(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrsetzone(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrtimer(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTIME b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrtlctldata(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTLCTL a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrtofs(ushort FlibHndl, short a, short b, short c, int d);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrtofsr(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTO_1_1 b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrtofsr(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTO_1_2 b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrtofsr(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTO_1_3 b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrtofsr(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTO_2 b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrtofsr(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTO_3 b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrtool(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTLMNG b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrtool_f2(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTLMNG_F2_data b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrtool2(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IDBTLM b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrtooldata(ushort FlibHndl, short a, ref short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTLDT c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrtoolzone(ushort FlibHndl, short a, ref short b, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBTLZN c);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrunsolicprm(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBUNSOLIC b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrwaveprm(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBWAVE a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrwaveprm2(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBWVPRM a);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrwkcdsfms(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBWCSF b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrwkcdshft(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBWCSF b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrzofs(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBZOFS b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_wrzofsr(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBZOR b);
        [DllImport("FWLIB32.dll")]
        public static extern short cnc_zofs_rnge(ushort FlibHndl, short a, short b, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBDATRNG c);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_cancel(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_checkhdd(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_chghdddir(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_copyhddfile(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, [In, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_delhdddir(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_delhddfile(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_delhostfile(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, int b);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_formathdd(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_gethostfile(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, [In, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_ldelhddfile(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_lgethostfile(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_lputhddfile(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_makehdddir(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_mgethostfile(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_mputhddfile(ushort hLib, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_puthddfile(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, [In, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_rddnchddfile(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_rddnchostfile(ushort FlibHndl, out short a, [Out, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_rdfile(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, [In, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_rdhdddir(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, int b, out short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBHDDDIR d);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_rdhddinfo(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBHDDINF a);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_rdhostdir(ushort FlibHndl, short a, int b, out short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBHOSTDIR d, int e);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_rdhostdir2(ushort FlibHndl, short a, int b, out short c, out int d, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBHOSTDIR e, int f);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_rdhostinfo(ushort FlibHndl, out int a, int b);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_rdhostno(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_rdm198hdddir(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_rdm198host(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_rdmntinfo(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.DSMNTINFO b);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_rdmode(ushort FlibHndl, ref short a);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_rdncfile(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_rdncfile2(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_rdresult(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_renhddfile(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, [In, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_searchresult(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_searchword(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_wrdnchddfile(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_wrdnchostfile(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_wrfile(ushort FlibHndl, [In, MarshalAs(UnmanagedType.AsAny)] object a, [In, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_wrhostno(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_wrm198hdddir(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_wrm198host(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_wrmode(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short ds_wrncfile(ushort FlibHndl, short a, int b);
        [DllImport("FWLIB32.dll")]
        public static extern short etb_rderrmsg(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBETMSG b);
        [DllImport("FWLIB32.dll")]
        public static extern short etb_rdparam(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBETP_ETB b);
        [DllImport("FWLIB32.dll")]
        public static extern short etb_rdparam(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBETP_FTP b);
        [DllImport("FWLIB32.dll")]
        public static extern short etb_rdparam(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBETP_HOST b);
        [DllImport("FWLIB32.dll")]
        public static extern short etb_rdparam(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBETP_TCP b);
        [DllImport("FWLIB32.dll")]
        public static extern short etb_wrparam(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBETP_FTP a);
        [DllImport("FWLIB32.dll")]
        public static extern short etb_wrparam(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBETP_HOST a);
        [DllImport("FWLIB32.dll")]
        public static extern short etb_wrparam(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBETP_TCP a);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_crdmsg(ushort FlibHndl, ref short a, [Out, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_cwrmsg(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_get_current_pmc_unit(ushort FlibHndl, ref int a);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_get_number_of_pmc(ushort FlibHndl, ref int a);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_get_pmc_unit_types(ushort FlibHndl, int[] a, ref int b);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_get_timer_type(ushort FlibHndl, ushort a, ushort b, ref short c);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_getdtailerr(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBPMCERR a);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_kpmsiz(ushort FlibHndl, out uint a);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_prfrdallcadr(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPRFADR b);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_prfrdbusprm(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBBUSPRM a);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_prfrdconfig(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBPRFCNF a);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_prfrddido(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBDIDO b);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_prfrdindiadr(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBINDEADR a);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_prfrdopmode(ushort FlibHndl, ref short a);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_prfrdslvaddr(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBSLVADR a);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_prfrdslvid(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBSLVID b);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_prfrdslvprm(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBSLVPRM b);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_prfrdslvprm(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBSLVPRM2 b);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_prfrdslvprm2(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBSLVPRM3 b);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_prfrdslvstat(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBSLVST a);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_prfwrallcadr(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPRFADR b);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_prfwrbusprm(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBBUSPRM a);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_prfwrdido(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBDIDO b);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_prfwrindiadr(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBINDEADR a);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_prfwropmode(ushort FlibHndl, short a, ref short b);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_prfwrslvaddr(ushort FlibHndl, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBSLVADR a);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_prfwrslvid(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBSLVID b);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_prfwrslvprm(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBSLVPRM b);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_prfwrslvprm(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBSLVPRM2 b);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_prfwrslvprm2(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBSLVPRM3 b);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_rdalmmsg(ushort FlibHndl, short a, ref short b, out short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBPMCALM d);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_rdcntldata(ushort FlibHndl, short a, short b, short c, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPMCCNTL d);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_rdcntlgrp(ushort FlibHndl, out short a);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_rdkpm(ushort FlibHndl, uint a, [Out, MarshalAs(UnmanagedType.AsAny)] object b, ushort c);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_rdkpm2(ushort FlibHndl, uint a, [Out, MarshalAs(UnmanagedType.AsAny)] object b, uint c);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_rdmsg(ushort FlibHndl, ref short a, [Out, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_rdpmcaddr(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBPMCADR a);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_rdpmcinfo(ushort FlibHndl, short a, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBPMCINF b);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_rdpmcmem(ushort FlibHndl, short a, int b, int c, [Out, MarshalAs(UnmanagedType.AsAny)] object d);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_rdpmcparam(ushort FlibHndl, ref int a, [Out, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_rdpmcrng(ushort FlibHndl, short a, short b, ushort c, ushort d, ushort e, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPMC0 f);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_rdpmcrng(ushort FlibHndl, short a, short b, ushort c, ushort d, ushort e, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPMC1 f);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_rdpmcrng(ushort FlibHndl, short a, short b, ushort c, ushort d, ushort e, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPMC2 f);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_rdpmcrng_ext(ushort FlibHndl, short a, [In, Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPMCEXT b);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_rdpmcsemem(ushort FlibHndl, short a, int b, int c, [Out, MarshalAs(UnmanagedType.AsAny)] object d);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_rdpmctitle(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.ODBPMCTITLE a);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_rdprmend(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_rdprmstart(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_select_pmc_unit(ushort FlibHndl, int a);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_set_timer_type(ushort FlibHndl, ushort a, ushort b, ref short c);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_wrcntldata(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPMCCNTL b);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_wrcntlgrp(ushort FlibHndl, short a);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_wriolinkdat(ushort FlibHndl, uint a, [In, MarshalAs(UnmanagedType.AsAny)] object b, uint c);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_wrkpm(ushort FlibHndl, uint a, [In, MarshalAs(UnmanagedType.AsAny)] object b, ushort c);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_wrkpm2(ushort FlibHndl, uint a, [In, MarshalAs(UnmanagedType.AsAny)] object b, uint c);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_wrmsg(ushort FlibHndl, short a, [In, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_wrpmcmem(ushort FlibHndl, short a, int b, int c, [In, MarshalAs(UnmanagedType.AsAny)] object d);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_wrpmcparam(ushort FlibHndl, ref int a, [In, MarshalAs(UnmanagedType.AsAny)] object b);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_wrpmcrng(ushort FlibHndl, ushort a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPMC0 b);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_wrpmcrng(ushort FlibHndl, ushort a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPMC1 b);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_wrpmcrng(ushort FlibHndl, ushort a, [In, MarshalAs(UnmanagedType.LPStruct)] FocasLibBase.IODBPMC2 b);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_wrpmcsemem(ushort FlibHndl, short a, int b, int c, [In, MarshalAs(UnmanagedType.AsAny)] object d);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_wrprmend(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short pmc_wrprmstart(ushort FlibHndl);        

        [DllImport("FWLIB32.dll")]
        public static extern short  cnc_rdcncid( ushort FlibHndl, uint [] cncid);

         [DllImport("FWLIB32.dll")]
        public static extern short cnc_exeprgname2(ushort FlibHndl, [Out, MarshalAs(UnmanagedType.AsAny)] object b);
        
    }
}


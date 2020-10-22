using System;
using System.Collections.Generic;
using System.Text;
using FocasLib;
using DTO;
using System.Diagnostics;
using System.IO;
using System.Threading;
namespace FocasLibrary
{
    public static partial class FocasData
    {
        //cnc_rdprogdir3
        public static List<ProgramDTO> ReadAllPrograms(string ipAddress, ushort portNo,short programDetailType=2)
        {
            List<ProgramDTO> programs = new List<ProgramDTO>();           
            int topProgram = 0;
            short prgromsToRead = 10;
            short ret = 0;
            ushort focasLibHandle = 0;
            ret = FocasLib.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
            if (ret != 0)
            {
                Logger.WriteErrorLog("cnc_allclibhndl3() failed. return value is = " + ret);
                return null;
            }
            ret = 0;
            while (prgromsToRead >= 10)
            {
                FocasLibBase.PRGDIR3 d = new FocasLibBase.PRGDIR3();
                try
                {
                    ret = FocasLib.cnc_rdprogdir3(focasLibHandle, programDetailType, ref topProgram, ref prgromsToRead, d);
                }
                catch (Exception ex)
                {
                    Logger.WriteErrorLog(ex.ToString());
                }
                if (ret != 0)
                {
                    Logger.WriteErrorLog("cnc_rdprogdir3() failed. return value is = " + ret);
                    break;
                }
                if (prgromsToRead == 0)
                {
                    Logger.WriteErrorLog("No programs found.");
                    break;
                }

                if (prgromsToRead >= 1) get_program(programs, d.dir1);
                if (prgromsToRead >= 2) get_program(programs, d.dir2);
                if (prgromsToRead >= 3) get_program(programs, d.dir3);
                if (prgromsToRead >= 4) get_program(programs, d.dir4);
                if (prgromsToRead >= 5) get_program(programs, d.dir5);
                if (prgromsToRead >= 6) get_program(programs, d.dir6);
                if (prgromsToRead >= 7) get_program(programs, d.dir7);
                if (prgromsToRead >= 8) get_program(programs, d.dir8);
                if (prgromsToRead >= 9) get_program(programs, d.dir9);
                if (prgromsToRead >= 10) get_program(programs, d.dir10);

                topProgram = programs[programs.Count - 1].ProgramNo + 1;

            }
            return programs;
        }

        public static ProgramDTO ReadOneProgram(ushort focasLibHandle, int progromToRead, short programDetailType = 1)
        {
            ProgramDTO program = new ProgramDTO();
            int topProgram = progromToRead;
            short prgromsToRead = 1;
            short ret = 0;
            ret = 0;
            FocasLibBase.PRGDIR3 d = new FocasLibBase.PRGDIR3();
            try
            {
                ret = FocasLib.cnc_rdprogdir3(focasLibHandle, programDetailType, ref topProgram, ref prgromsToRead, d);               

                if (ret != 0)
                {
                    Logger.WriteErrorLog("cnc_rdprogdir3() failed. return value is = " + ret);
                }
                else
                {
                    if (prgromsToRead > 0)
                    {
                        if (topProgram == progromToRead)
                        {
                            program.ProgramNo = d.dir1.number;
                            program.Comment = d.dir1.comment;
                        }
                    }
                }             
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            } 
            return program;
        }

        private static void get_program(List<ProgramDTO> ls, FocasLibBase.PRGDIR3_data dir)
        {
            ProgramDTO pDto = new ProgramDTO();
            pDto.ProgramNo = dir.number;
            pDto.ProgramLenght = dir.length;
            pDto.Comment = dir.comment;
            pDto.ModifiedDate = new DateTime(dir.mdate.year < 2000 ? dir.mdate.year + 2000 : dir.mdate.year, dir.mdate.month, dir.mdate.day, dir.mdate.hour, dir.mdate.minute, 00);           
            ls.Add(pDto);
        }

        //to download use below functions
        //cnc_upstart3,cnc_upload3, cnc_upend3
        public static bool DownloadProgram_OLD(string ipAddress, ushort portNo, int programNo)
        {
            short ret = -20;
            int bufsize = 1280;
            int dataLength = 0;
            ushort focasLibHandle = 0;
            ret = FocasLib.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
            if (ret != 0)
            {
                Logger.WriteErrorLog("cnc_allclibhndl3() failed. return value is = " + ret);
                return false;
            }
            ret = FocasLib.cnc_upstart3(focasLibHandle, 0, programNo, programNo);
            if (ret != 0)
            {
                Logger.WriteErrorLog("cnc_upstart3() failed. return value is = " + ret);
                return false;
            }
            StringBuilder programStr = new StringBuilder(bufsize*10);           
            bool endFound = false;
            do
            {
                char[] buf = new char[bufsize + 1];
                dataLength = bufsize;
                ret = FocasLib.cnc_upload3(focasLibHandle, ref dataLength, buf);

                if (ret == 10) // if buffer is empty retry
                {
                    Thread.Sleep(400);
                    continue;
                   
                }
                if (ret == -2) //if buffer is in reset mode
                {
                    Logger.WriteErrorLog("cnc_upload3() failed. return value is = " + ret);
                    break;
                }
                if (ret != 0)
                {
                    Logger.WriteErrorLog("cnc_upload3() failed. return value is = " + ret);
                    break;
                }

                char[] tempBuf = new char[dataLength];
                Array.Copy(buf, tempBuf, dataLength);
                programStr.Append(tempBuf);
                if (buf[dataLength - 1] == '%')
                {
                    endFound = true;
                    break;
                   
                }
                Thread.Sleep(600);

            } while (endFound == false);
            File.WriteAllText("D:\\O9999.txt", programStr.Replace("\r", "").ToString());
            ret = FocasLib.cnc_upend3(focasLibHandle);
            if (ret != 0)
            {
                Logger.WriteErrorLog("cnc_upend3() failed. return value is = " + ret);               
            }

            return true;
        }

        //to sent to cnc
        //cnc_dwnstart3, cnc_download3, cnc_dwnend3
        public static bool UploadProgram(string ipAddress, ushort portNo, string program)
        {
            short ret = -20;
            int len, n;
            program = File.ReadAllText("D:\\O12345.txt");

            ushort focasLibHandle = 0;
            ret = FocasLib.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
            if (ret != 0)
            {
                Logger.WriteErrorLog("cnc_allclibhndl3() failed. return value is = " + ret);
                return false;
            }
            ret = FocasLib.cnc_dwnstart3(focasLibHandle, 0);
            if (ret != 0)
            {
                Logger.WriteErrorLog("cnc_dwnstart3() failed. return value is = " + ret);
                return false;
            }
            if (ret == FocasLib.EW_OK)
            {
                len = program.Length;
                while (len > 0)
                {
                    n = program.Length;
                    ret = FocasLib.cnc_download3(focasLibHandle, ref n, program);
                   
                    if (ret == 10) //if buffer is empty
                    {
                        continue;
                    }
                    if (ret == -2) // if buffer in reset mode
                    {
                        break;
                    }
                    if (ret != FocasLib.EW_OK)
                    {
                        Logger.WriteErrorLog("cnc_download3() failed. return value is = " + ret);
                        break;
                    }
                   
                    if (ret == FocasLib.EW_OK)
                    {
                        //program += n;
                        //len -= n;

                        program = program.Substring(n,len-n);
                        len -= n;
                    } 
                }

                ret = FocasLib.cnc_dwnend(focasLibHandle);
                if (ret != FocasLib.EW_OK)
                {
                    Logger.WriteErrorLog("cnc_dwnend() failed. return value is = " + ret);                    
                }                        
            }
            return true;      
        }

        //to search        
        /*
          The behavior of this function depends on the CNC mode. 
          EDIT, MEM mode : foreground search 
          Other mode : background search 
         */
        public static bool SearchPrograms(string ipAddress, ushort portNo,short programNumber)
        {
            bool isFound = false;
            focas_ret ret;
            int ret1 = -1;
            ushort focasLibHandle = 0;
            ret1 = FocasLib.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
            if (ret1 != 0)
            {
                Logger.WriteErrorLog("cnc_allclibhndl3() failed. return value is = " + ret1);
                return false;
            }
            short retShort = FocasLib.cnc_search(focasLibHandle, programNumber);
            Enum.TryParse<focas_ret>(retShort.ToString(), out ret);
            switch (ret)
            {
                case focas_ret.EW_OK:
                    Logger.WriteDebugLog("PROGRAM " + programNumber + "  has been searched.");
                    isFound = true;
                    break;
                case focas_ret.EW_DATA:
                    Logger.WriteDebugLog("PROGRAM " + programNumber + " doesn't exist.");
                    break;
                case focas_ret.EW_PROT:
                    Logger.WriteDebugLog("Program " + programNumber + " is PROTECTED.");
                    break;
                case focas_ret.EW_BUSY:
                    Logger.WriteDebugLog("CNC is busy. Program " + programNumber + " search REJECTED.");
                    break;                
            }
            return isFound;
        }

        //to delete EW_PASSWD
        public static bool DeletePrograms(string ipAddress, ushort portNo, short programNumber)
        {
            bool isdeleted = false;
            focas_ret ret;
            int ret1 = -1;
            ushort focasLibHandle = 0;
            ret1 = FocasLib.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
            if (ret1 != 0)
            {
                Logger.WriteErrorLog("cnc_allclibhndl3() failed. return value is = " + ret1);
                return false;
            }
            short retShort = FocasLib.cnc_delete(focasLibHandle, programNumber);
            Enum.TryParse<focas_ret>(retShort.ToString(), out ret);
            switch (ret)
            {
                case focas_ret.EW_OK:
                    Logger.WriteDebugLog("PROGRAM" + programNumber + " has been deleted.");
                    isdeleted = true;
                    break;
                case focas_ret.EW_DATA:
                    Logger.WriteDebugLog("PROGRAM " + programNumber + " doesn't exist.");
                    break;
                case focas_ret.EW_PROT:
                    Logger.WriteDebugLog("Program " + programNumber + " is PROTECTED.");
                    break;
                case focas_ret.EW_BUSY:
                    Logger.WriteDebugLog("CNC is busy. Program " + programNumber + "  delete REJECTED.");
                    break;
                case focas_ret.EW_PASSWD:
                    Logger.WriteDebugLog("Specified program number cannot be deleted because the data is protected.");
                    break;
            }
            return isdeleted;
        }

        //running program number
        //cnc_rdprgnum       
        public static short ReadRunningProgramNumber(ushort handle)
        {
            FocasLibBase.ODBPRO odbpro = new FocasLibBase.ODBPRO();
            FocasLib.cnc_rdprgnum(handle, odbpro);
            return odbpro.mdata;            
        }

        public static bool DownloadOperationHistory_TEST(string ipAddress, ushort portNo, int programNo)
        {
            short ret = -20;
            int bufsize = 2560;
            int dataLength = 0;
            ushort focasLibHandle = 0;
            ret = FocasLib.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
            if (ret != 0)
            {
                Logger.WriteErrorLog("cnc_allclibhndl3() failed. return value is = " + ret);
                return false;
            }
            ret = FocasLib.cnc_upstart3(focasLibHandle, 7, programNo, programNo);
            if (ret != 0)
            {
                Logger.WriteErrorLog("cnc_upstart3() failed. return value is = " + ret);
                return false;
            }
            StringBuilder programStr = new StringBuilder(bufsize * 10);
            bool endFound = false;
            do
            {
                char[] buf = new char[bufsize + 1];
                dataLength = bufsize;
                ret = FocasLib.cnc_upload3(focasLibHandle, ref dataLength, buf);

                if (ret == 10) // if buffer is empty retry
                {
                    Thread.Sleep(400);
                    continue;
                }
                if (ret == -2) //if buffer is in reset mode
                {
                    Logger.WriteErrorLog("cnc_upload3() failed. return value is = " + ret);
                    break;
                }
                if (ret != 0)
                {
                    Logger.WriteErrorLog("cnc_upload3() failed. return value is = " + ret);
                    break;
                }

                char[] tempBuf = new char[dataLength];
                Array.Copy(buf, tempBuf, dataLength);
                programStr.Append(tempBuf);
                if (buf[dataLength - 1] == '%')
                {
                    endFound = true;
                    break;
                }
                Thread.Sleep(100);

            } while (endFound == false);
            ret = FocasLib.cnc_upend3(focasLibHandle);
            if (ret != 0)
            {
                Logger.WriteErrorLog("cnc_upend3() failed. return value is = " + ret);
            }
            File.WriteAllText("D:\\O9999.txt", programStr.Replace("\r", "").ToString());           
            return true;
        }

        //to download use below functions
        //cnc_upstart3,cnc_upload3, cnc_upend3
        public static bool DownloadOperationHistory(string ipAddress, ushort portNo, string file)
        {            
            short ret = -20;
            int bufsize = 1280;
            int dataLength = 0;
            int programNo = 0;
            ushort focasLibHandle = 0;
            ret = FocasLib.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
            if (ret != 0)
            {
                Logger.WriteErrorLog("cnc_allclibhndl3() failed. return value is = " + ret);
                Logger.WriteErrorLog("Not able to connect to CNC machine. Please check the network connection and try again");
                return false;
            }

            int count = 0;
            ret = short.MinValue;            
            while (ret != 0 && count < 20)
            {
                ret = FocasLib.cnc_upstart3(focasLibHandle, 7, programNo, programNo);                
                count++;
                if (ret == -1) Thread.Sleep(400);
            }
            if (ret == -1)
            {
                Logger.WriteErrorLog("CNC is busy. Please try after some time.");
                //return false;
            }

            if (ret != 0)
            {
                Logger.WriteErrorLog("cnc_upstart3() failed. return value is = " + ret);
                FocasLib.cnc_freelibhndl(focasLibHandle);
                return false;
            }
            StringBuilder programStr = new StringBuilder(bufsize * 10);
            bool endFound = false;
            do
            {
                char[] buf = new char[bufsize + 1];
                dataLength = bufsize;
                ret = FocasLib.cnc_upload3(focasLibHandle, ref dataLength, buf);

                if (ret == 10) // if buffer is empty retry
                {
                    Thread.Sleep(400);
                    continue;

                }
                if (ret == -2) //if buffer is in reset mode
                {
                    Logger.WriteErrorLog("cnc_upload3() failed. return value is = " + ret);
                    break;
                }

                if (ret == -16)
                {
                    Logger.WriteErrorLog("cnc_upload3() failed. return value is = " + ret);
                    ret = FocasLib.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
                    if (ret != 0)
                    {
                        Logger.WriteErrorLog("cnc_upload3() failed.return value is = " + ret);
                        break;
                    }
                    continue;                   
                }
                if (ret != 0)
                {
                    Logger.WriteErrorLog("cnc_upload3() failed. return value is = " + ret);
                    break;
                }

                char[] tempBuf = new char[dataLength];
                Array.Copy(buf, tempBuf, dataLength);
                programStr.Append(tempBuf);
                if (buf[dataLength - 1] == '%')
                {
                    endFound = true;
                    break;
                }
                Thread.Sleep(100);

            } while (endFound == false);

            ret = FocasLib.cnc_upend3(focasLibHandle);
            if (ret != 0)
            {
                Logger.WriteErrorLog("cnc_upend3() failed. return value is = " + ret);
                //return false;
            }
            if (string.IsNullOrEmpty(Convert.ToString(programStr)))
            {
                Logger.WriteDebugLog("program is empty.");
                FocasLib.cnc_freelibhndl(focasLibHandle);
                return false;
            }
            programStr.Replace("\r", "").Replace("\n", "\r\n");

            try
            {
                File.WriteAllText(file, programStr.ToString());
            }
            catch (Exception ex)
            {
                Logger.WriteDebugLog(ex.ToString());
            }
            FocasLib.cnc_freelibhndl(focasLibHandle);
            return true;
        }


        //to download use below functions
        //cnc_upstart3,cnc_upload3, cnc_upend3
        public static string DownloadProgram(string ipAddress, ushort portNo, int programNo, out bool result)
        {
            result = false;
            short ret = -20;
            int bufsize = 1280;
            int dataLength = 0;
            ushort focasLibHandle = 0;
            string programDownloaded = string.Empty;
            ret = FocasLib.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
            if (ret != 0)
            {
                Logger.WriteErrorLog("cnc_allclibhndl3() failed. return value is = " + ret + "Not able to connect to CNC machine. Please check the network connection and try again");                
                return string.Empty;
            }
            int count = 0;
            ret = short.MinValue;

            while (ret != 0 && count < 10)
            {
                ret = FocasLib.cnc_upstart3(focasLibHandle, 0, programNo, programNo);
                count++;
                if (ret == -1) Thread.Sleep(400);
            }
            if (ret == -1)
            {
                Logger.WriteDebugLog("CNC is busy. Please try after some time.");
                FocasLib.cnc_freelibhndl(focasLibHandle);
                return string.Empty;
            }

            if (ret != 0)
            {
                Logger.WriteErrorLog("cnc_upstart3() failed. return value is = " + ret);
                FocasLib.cnc_freelibhndl(focasLibHandle);               
                return string.Empty;
            }
            StringBuilder programStr = new StringBuilder(bufsize * 10);
            bool endFound = false;
            do
            {
                char[] buf = new char[bufsize + 1];
                dataLength = bufsize;
                ret = FocasLib.cnc_upload3(focasLibHandle, ref dataLength, buf);

                if (ret == 10) // if buffer is empty retry
                {
                    Thread.Sleep(400);
                    continue;

                }
                if (ret == -2) //if buffer is in reset mode Write protected on CNC side
                {
                    Logger.WriteErrorLog("Reset or stop request. The data to read is nothing. cnc_upload3() failed. return value is = " + ret);                   
                    break;
                }
                if (ret == 7) //if buffer is in reset mode Write protected on CNC side
                {
                    Logger.WriteErrorLog("Write protected on CNC side.. cnc_upload3() failed. return value is = " + ret);                   
                    break;
                }
                if (ret != 0)
                {
                    Logger.WriteErrorLog("cnc_upload3() failed. return value is = " + ret);                   
                    break;
                }

                char[] tempBuf = new char[dataLength];
                Array.Copy(buf, tempBuf, dataLength);
                programStr.Append(tempBuf);
                if (buf[dataLength - 1] == '%')
                {
                    result  = endFound = true;                    
                    break;
                }
                Thread.Sleep(600);

            } while (endFound == false);
            if (ret == 7)
            {
                FocasLib.cnc_freelibhndl(focasLibHandle);
                return string.Empty;
            }

            ret = FocasLib.cnc_upend3(focasLibHandle);
            if (ret != 0)
            {
                Logger.WriteErrorLog("cnc_upend3() failed. return value is = " + ret);               
            }
            if (string.IsNullOrEmpty(Convert.ToString(programStr)))
            {
                Logger.WriteDebugLog("program is empty.");
                FocasLib.cnc_freelibhndl(focasLibHandle);
                return string.Empty;
            }
            programStr.Replace("\r", "").Replace("\n", "\r\n");
            result = true;           
            FocasLib.cnc_freelibhndl(focasLibHandle);
            return programStr.ToString();          
        }

        private static string GetVersion(string folder, string file, StringBuilder program)//returns complete filepath(with version)//
        {
            bool matching = false;
            int i = 1;
            file = file + ".txt";

            if (!File.Exists(Path.Combine(folder, file)))//if 1000 not exists??
            {
                return Path.Combine(folder, file);

            }
            else
            {
                matching = CompareContents(Path.Combine(folder, file), program);
                if (matching == true)
                {
                    return "Same version of program exists on the computer system as " + Path.GetFileNameWithoutExtension(Path.Combine(folder, file));
                }

                string pwithext = "";
                string path_wtout_exten = Path.GetFileNameWithoutExtension(Path.Combine(folder, file));//
                while (true)
                {
                    pwithext = Path.Combine(folder, path_wtout_exten + "." + i.ToString() + ".txt");
                    if (!File.Exists(pwithext))//here we can comapre with folder+file contents also//
                    {
                        return pwithext;
                    }
                    else//if file exists
                    {

                        matching = CompareContents(pwithext, program);//compare with downloaded file//
                        if (matching == true)
                        {
                            return "Same Program EXISTS as " + pwithext;

                        }
                        i++;
                        if (i >= 1000)
                            return "ERROR";
                    }
                }
            }
        }

        private static bool CompareContents(string pwithext, StringBuilder programRecieved)
        {
            //TODO
            string str1 = File.ReadAllText(pwithext);
            if (str1 == programRecieved.ToString())
                return true;

            return false;
        }


        public static void UpdateOperatinHistoryMacroLocation(string ipAddress, ushort portNo,short macroLocation, int value)
        {
            short ret = -20;           
            ushort focasLibHandle = 0;
            ret = FocasLib.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
            if (ret != 0)
            {
                Logger.WriteErrorLog("cnc_allclibhndl3() failed. return value is = " + ret);
                return;
            }
            FocasData.WriteMacro(focasLibHandle, macroLocation, value);
            FocasLib.cnc_freelibhndl(focasLibHandle);           
        }

        public static int ReadOperatinHistoryDPrintLocation(string ipAddress, ushort portNo, short macroLocation)
        {
            short ret = -20;
            ushort focasLibHandle = 0;
            int dprintValue = 0;
            ret = FocasLib.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
            if (ret != 0)
            {
                Logger.WriteErrorLog("cnc_allclibhndl3() failed. return value is = " + ret);
                return dprintValue;
            }
            dprintValue = FocasData.ReadMacro(focasLibHandle, macroLocation);
            FocasLib.cnc_freelibhndl(focasLibHandle);
            return dprintValue;
        }

        public static string ReadFullProgramPathRunningProgram(ushort handle)
        {
            char[] buf = new char[256 + 1];
            short ret = FocasLib.cnc_pdf_rdmain(handle, buf);
            StringBuilder str = new StringBuilder();
            str.Append(buf);
           return str.ToString().Trim('\0');
        }

        public static string ReadFullProgramPathRunningProgram(string ipAddress, ushort portNo)
        {
            short ret = -20;
            ushort focasLibHandle = 0;
            char[] buf = new char[256 + 1];

            ret = FocasLib.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
            if (ret != 0)
            {
                Logger.WriteErrorLog("cnc_allclibhndl3() failed. return value is = " + ret);
                return string.Empty;
            }

            ret = FocasLib.cnc_pdf_rdmain(focasLibHandle, buf);
            if (ret != 0) return string.Empty;
            StringBuilder str = new StringBuilder();
            str.Append(buf);
            return str.ToString().Trim('\0');
        }

        public static string DownloadProgram(string ipAddress, ushort portNo, int programNo, out bool result, string folderPath, bool isProgramFolderSupports)
        {
            result = false;
            short ret = -20;
            int bufsize = 1280;
            int dataLength = 0;
            ushort focasLibHandle = 0;
            string programDownloaded = string.Empty;
            string programNoStr = "";

            if (isProgramFolderSupports)
            {
                programNoStr = "O" + programNo;
            }


            ret = FocasLib.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
            if (ret != 0)
            {
                Logger.WriteErrorLog("cnc_allclibhndl3() failed. return value is = " + ret + "Not able to connect to CNC machine. Please check the network connection and try again");
                return string.Empty;
            }
            int count = 0;
            ret = short.MinValue;

            while (ret != 0 && count < 60)
            {
                //ret = FocasLib.cnc_upstart3(focasLibHandle, 0, programNo, programNo);
                if (isProgramFolderSupports)
                {
                    ret = FocasLib.cnc_upstart4(focasLibHandle, 0, Path.GetDirectoryName(folderPath).Replace("\\", "/") + "/" + programNoStr);
                }
                else
                {
                    ret = FocasLib.cnc_upstart3(focasLibHandle, 0, programNo, programNo);
                }
                count++;
                if (ret == -1)
                {
                    Thread.Sleep(1000);
                    ret = FocasLib.cnc_upend4(focasLibHandle);
                }
            }
            if (ret == -1)
            {
                Logger.WriteDebugLog("CNC is busy. Please try after some time.");
                FocasLib.cnc_freelibhndl(focasLibHandle);
                return string.Empty;
            }

            if (ret != 0)
            {
                Logger.WriteErrorLog("cnc_upstart3() failed. return value is = " + ret);
                FocasLib.cnc_freelibhndl(focasLibHandle);
                return string.Empty;
            }
            StringBuilder programStr = new StringBuilder(bufsize * 10);
            bool endFound = false;
            do
            {
                char[] buf = new char[bufsize + 1];
                dataLength = bufsize;
                //ret = FocasLib.cnc_upload3(focasLibHandle, ref dataLength, buf);
                if (isProgramFolderSupports)
                {
                    ret = FocasLib.cnc_upload4(focasLibHandle, ref dataLength, buf);
                }
                else
                {
                    ret = FocasLib.cnc_upload3(focasLibHandle, ref dataLength, buf);
                }

                if (ret == 10) // if buffer is empty retry
                {
                    Thread.Sleep(400);
                    continue;

                }
                if (ret == -2) //if buffer is in reset mode Write protected on CNC side
                {
                    Logger.WriteErrorLog("Reset or stop request. The data to read is nothing. cnc_upload3() failed. return value is = " + ret);
                    break;
                }
                if (ret == 7) //if buffer is in reset mode Write protected on CNC side
                {
                    Logger.WriteErrorLog("Write protected on CNC side.. cnc_upload3() failed. return value is = " + ret);
                    break;
                }
                if (ret != 0)
                {
                    Logger.WriteErrorLog("cnc_upload3() failed. return value is = " + ret);
                    break;
                }

                char[] tempBuf = new char[dataLength];
                Array.Copy(buf, tempBuf, dataLength);
                programStr.Append(tempBuf);
                if (buf[dataLength - 1] == '%')
                {
                    result = endFound = true;
                    break;
                }
                Thread.Sleep(600);

            } while (endFound == false);
            if (ret == 7)
            {
                FocasLib.cnc_freelibhndl(focasLibHandle);
                return string.Empty;
            }

            //ret = FocasLib.cnc_upend3(focasLibHandle);
            if (isProgramFolderSupports)
            {
                ret = FocasLib.cnc_upend4(focasLibHandle);
            }
            else
            {
                ret = FocasLib.cnc_upend3(focasLibHandle);
            }
            if (ret != 0)
            {
                Logger.WriteErrorLog("cnc_upend3() failed. return value is = " + ret);
            }
            if (string.IsNullOrEmpty(Convert.ToString(programStr)))
            {
                Logger.WriteDebugLog("program is empty.");
                FocasLib.cnc_freelibhndl(focasLibHandle);
                return string.Empty;
            }
            programStr.Replace("\r", "").Replace("\n", "\r\n");
            result = true;
            FocasLib.cnc_freelibhndl(focasLibHandle);
            return programStr.ToString();
        }
    
    }
}

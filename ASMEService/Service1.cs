using System;
using System.ServiceProcess;
using Npgsql;
using ASMeSDK_CSharp;
using System.Collections.Generic;
using System.Configuration;

namespace ASMEService
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
            _shouldLog = string.Equals(ConfigurationManager.AppSettings["log"], "true") ? true : false;
            _timerInterval = Convert.ToDouble(ConfigurationManager.AppSettings["timer"]);

            // _ipCheck = 0;
        }

        System.Timers.Timer timer;
        static int OpenControllerStop = 0;
        IntPtr m_hController = IntPtr.Zero;
        private readonly bool _shouldLog;
        private readonly double _timerInterval;

        protected override void OnStart(string[] args)
        {
            WriteToFile("Service is started at " + DateTime.Now);
            timer = new System.Timers.Timer(_timerInterval);
            timer.Elapsed += new System.Timers.ElapsedEventHandler(OnElapsedTime);
            timer.Enabled = true;
        }

        private void OnElapsedTime(object source, System.Timers.ElapsedEventArgs e)
        {
            //WriteToFile("Service recall at " + DateTime.Now);
            if (_getRecords(true) > 0)
                _getRecords(false);
        }

        protected override void OnStop()
        {
            WriteToFile("Service is stopped at " + DateTime.Now);
        }

        private static uint StrIpToUint(string strIPAddress)
        {
            if (string.IsNullOrEmpty(strIPAddress)) return 0;

            string[] ipSplit = strIPAddress.Split('.');

            return (uint.Parse(ipSplit[3])) | (uint.Parse(ipSplit[2]) << 8) |
                (uint.Parse(ipSplit[1]) << 16) | (uint.Parse(ipSplit[0]) << 24);
        }

        private int _getRecords(bool dev)
        {
            try
            {
                var devices = dev ? GetDevices("select device_ip, device_door from devices where device_status = 'enter'"):
                                    GetDevices("select device_ip, device_door from devices where device_status = 'exit'");
                
                if (devices.Count < 1)
                {
                    WriteToFile("No device " + dev);
                    return -1;
                }

                foreach (var device in devices)
                {
                    asc_STU.LPAS_ME_COMM_ADDRESS address = new asc_STU.LPAS_ME_COMM_ADDRESS
                    {
                        nType = asc_STU.AS_ME_COMM_TYPE_IPV4,
                        IPV4 = new asc_STU.IPV4
                        {
                            dwIPAddress = StrIpToUint(device.Ip),
                            wServicePort = 50000
                        }
                    };
                    int nRes = asc_SDKAPI.AS_ME_OpenController((int)asc_STU.asc_device_type.AS_ME_TYPE_UNKOWN, ref address, 0,
                            asc_STU.AS_ME_NO_PASSWORD, ref OpenControllerStop, ref m_hController);
                    if (nRes < 0)
                    {
                        WriteToFile("AsmeDevice AS_ME_OpenController " + device.Ip + ": " + nRes);
                        return -1;
                    }

                    asc_STU.SYSTEMTIME st = new asc_STU.SYSTEMTIME
                    {
                        wYear = (ushort)DateTime.Now.Year,
                        wMonth = (ushort)DateTime.Now.Month,
                        wDay = (ushort)DateTime.Now.Day,

                        wHour = (ushort)DateTime.Now.Hour,
                        wMinute = (ushort)DateTime.Now.Minute,
                        wSecond = (ushort)DateTime.Now.Second,
                        wDayOfWeek = (ushort)DateTime.Now.DayOfWeek
                    };

                    nRes = asc_SDKAPI.AS_ME_SetTime(m_hController, ref st);
                    if (nRes < 0)
                        WriteToFile("Setting the Time filed " + nRes);

                    asc_SDKAPI.AS_ME_GetTime(m_hController, ref st);

                    asc_STU.LPAS_ME_RECORDS aRecords = new asc_STU.LPAS_ME_RECORDS(); 
                    string query = "", leave_time = "";

                    while (true)
                    {
                        nRes = asc_SDKAPI.AS_ME_GetEventRecord(m_hController, ref aRecords);
                        if (nRes < 0)
                        {
                            if (_shouldLog)
                            {
                                WriteToFile("AsmeDevice No records " + nRes);
                            }
                            break;
                        }

                        if (aRecords.nCount > 0)
                        {
                            for (int i = 0; i < aRecords.nCount; i++)
                            {
                                switch (aRecords.aEventRecord[i].nType)
                                {
                                    case 0: // legal card
                                        leave_time = string.Format(("{0:d}-{1:d2}-{2:d2} {3:d2}:{4:d2}:{5:d2}"),
                                        aRecords.aEventRecord[i].stStamp.wYear, aRecords.aEventRecord[i].stStamp.wMonth,
                                        aRecords.aEventRecord[i].stStamp.wDay, aRecords.aEventRecord[i].stStamp.wHour,
                                        aRecords.aEventRecord[i].stStamp.wMinute, aRecords.aEventRecord[i].stStamp.wSecond);

                                        query = (dev) ? "insert into reports (employeeid, door_name, status, kirish) " +
                                                "select employeeid, '" + device.Door + "', 0, '" + leave_time + "' from employee " +
                                                "where card = '" + aRecords.aEventRecord[i].dw64ID + "'" :
                                                "select update_reports((select employeeid from employee where card = '" + aRecords.aEventRecord[i].dw64ID +"'),'" + 
                                                device.Door + "','" + leave_time + "')";

                                        if (_shouldLog)
                                        {
                                            WriteToFile("Query " + query + "     \n" + "Temperature " + aRecords.aEventRecord[i].nBodyTemperature);
                                        }

                                        InsertDB(query);                                       

                                        break;

                                    case 2: // legal password
                                        leave_time = string.Format(("{0:d}-{1:d2}-{2:d2} {3:d2}:{4:d2}:{5:d2}"),
                                        aRecords.aEventRecord[i].stStamp.wYear, aRecords.aEventRecord[i].stStamp.wMonth,
                                        aRecords.aEventRecord[i].stStamp.wDay, aRecords.aEventRecord[i].stStamp.wHour,
                                        aRecords.aEventRecord[i].stStamp.wMinute, aRecords.aEventRecord[i].stStamp.wSecond);

                                        query = (dev) ? "insert into reports (employeeid, door_name, status, kirish) " +
                                            "select employeeid, '" + device.Door + "', 0, '" + leave_time + "' from employee " +
                                            "where card = '" + aRecords.aEventRecord[i].dw64ID + "'" :
                                            "select update_reports((select employeeid from employee where card = '" + aRecords.aEventRecord[i].dw64ID + "'),'" +
                                            device.Door + "','" + leave_time + "')";

                                        if (_shouldLog)
                                        {
                                            WriteToFile("Query " + query + "     \n" + "Temperature " + aRecords.aEventRecord[i].nBodyTemperature);
                                        }

                                        InsertDB(query);                                       

                                        break;

                                    case 13:
                                        leave_time = string.Format(("{0:d}-{1:d2}-{2:d2} {3:d2}:{4:d2}:{5:d2}"),
                                        aRecords.aEventRecord[i].stStamp.wYear, aRecords.aEventRecord[i].stStamp.wMonth,
                                        aRecords.aEventRecord[i].stStamp.wDay, aRecords.aEventRecord[i].stStamp.wHour,
                                        aRecords.aEventRecord[i].stStamp.wMinute, aRecords.aEventRecord[i].stStamp.wSecond);

                                        query = (dev) ? "insert into reports (employeeid, door_name, status, kirish) values ("
                                            + aRecords.aEventRecord[i].dw64ID + ",'" + device.Door + "', 0, '" + leave_time + "')" :
                                           "select update_reports(" + aRecords.aEventRecord[i].dw64ID + ",'" + device.Door + "','" +
                                           leave_time + "')";

                                        if (aRecords.aEventRecord[i].nBodyTemperature > 0)
                                            InsertDB("insert into temperature (employeeid, temperature, door, sana) values " +
                                                "(" + aRecords.aEventRecord[i].dw64ID + "," + aRecords.aEventRecord[i].nBodyTemperature +
                                                ",'" + device.Door + "','" + leave_time + "')");

                                        if (_shouldLog)
                                        {
                                            WriteToFile("Query " + query + "     \n" + "Temperature " + aRecords.aEventRecord[i].nBodyTemperature);
                                        }

                                        InsertDB(query);                                       

                                        break;

                                    case 11:
                                        leave_time = string.Format(("{0:d}-{1:d2}-{2:d2} {3:d2}:{4:d2}:{5:d2}"),
                                        aRecords.aEventRecord[i].stStamp.wYear, aRecords.aEventRecord[i].stStamp.wMonth,
                                        aRecords.aEventRecord[i].stStamp.wDay, aRecords.aEventRecord[i].stStamp.wHour,
                                        aRecords.aEventRecord[i].stStamp.wMinute, aRecords.aEventRecord[i].stStamp.wSecond);

                                        query = (dev) ? "insert into reports (employeeid, door_name, status, kirish) values ("
                                            + aRecords.aEventRecord[i].dw64ID + ",'" + device.Door + "', 0, '" + leave_time + "')" :
                                           "select update_reports(" + aRecords.aEventRecord[i].dw64ID + ",'" + device.Door + "','" +
                                           leave_time + "')";

                                        if (_shouldLog)
                                        {
                                            WriteToFile("Query " + query);
                                        }

                                        InsertDB(query);                                       

                                        break;
                                    default:
                                        if (_shouldLog)
                                        {
                                            WriteToFile("Illegal Face / Finger " + aRecords.aEventRecord[i].dw64ID);
                                        }
                                        break;
                                }
                            }
                        }
                        else
                        {
                            break;
                        }
                    }
                    //System.Runtime.InteropServices.Marshal.FreeHGlobal(infosIntptr);
                    //WriteToFile("Cleared device: " + device.Ip + "  dev:" + dev);
                    asc_SDKAPI.AS_ME_CloseController(m_hController);
                    m_hController = IntPtr.Zero;
                }
            }
            catch (Exception msg)
            {
                WriteToFile("Exception: " + msg.ToString());
                if(m_hController != IntPtr.Zero)
                {
                    asc_SDKAPI.AS_ME_CloseController(m_hController);
                    m_hController = IntPtr.Zero;
                }                
                return -1;
            }
            return 1;
        }

        public void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!System.IO.Directory.Exists(path))
            {
                System.IO.Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt";
            if (!System.IO.File.Exists(filepath))
            {
                // Create a file to write to.   
                using (System.IO.StreamWriter sw = System.IO.File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (System.IO.StreamWriter sw = System.IO.File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }

        private List<Device> GetDevices(string query)
        {
            var devices = new List<Device>();
            using (var conn = new NpgsqlConnection(Helper.CnnVal("DBConnection")))
            {
                using (var cmd = new NpgsqlCommand(query, conn))
                {
                    conn.Open();
                    var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        var device = new Device
                        {
                            Ip = reader["device_ip"].ToString(),
                            Door = reader["device_door"].ToString()
                        };
                        devices.Add(device);
                    }                    
                }
            }
            return devices;
        }

        private void InsertDB(string query)
        {
            using (var conn = new NpgsqlConnection(Helper.CnnVal("DBConnection")))
            {
                using (var cmd = new NpgsqlCommand(query, conn))
                {
                    try
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        WriteToFile($"DB: {ex.ToString()}");
                    }
                    finally
                    {
                        conn.Close();
                    }
                    
                }
            }
        }
    }
}

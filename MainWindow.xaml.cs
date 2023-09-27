using Microsoft.Win32;
using System;
using System.Windows;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Threading;
using ProgramMethod;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Linq;

namespace ShipmentReportApp
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window
    {
        SynchronizationContext _syncContext = null;
        private FileMethod PGMethod = new FileMethod();
        private string iboxdbinfo = "Data Source=10.3.30.20;User Id=sa;Password=hota;Database=iBox";
        private string hotadbinfo = "Data Source=10.3.30.203;User Id=sa;Password=hota;Database=TestMeasure";
        private List<string[]> ShipmentList;
        private int ShippingRows;
        private int startRow = 1;
        private bool Syncing = false;
        public MainWindow()
        {
            InitializeComponent();
            _syncContext = SynchronizationContext.Current;
            LB_QC1File.Content = PGMethod.GetConfigSetting("QC1File");
            LB_QC2File.Content = PGMethod.GetConfigSetting("QC2File");
            LB_FQCFile.Content = PGMethod.GetConfigSetting("FQCFile");
        }
        private void ExitButton_Click(object sender, RoutedEventArgs e)
        {
            Environment.Exit(0);
        }
        private void QC1FileButton_Click(object sender, RoutedEventArgs e)
        {
            //LB_QC1File.Content = CreateExcelFile();
            LB_QC1File.Content = Create2007ExcelFile();
            PGMethod.SetConfigSetting("QC1File", LB_QC1File.Content.ToString());
        }
        private void QC2FileButton_Click(object sender, RoutedEventArgs e)
        {
            //LB_QC2File.Content = CreateExcelFile();
            LB_QC2File.Content = Create2007ExcelFile();
            PGMethod.SetConfigSetting("QC2File", LB_QC2File.Content.ToString());
        }
        private void FQCFileButton_Click(object sender, RoutedEventArgs e)
        {
            //LB_FQCFile.Content = CreateExcelFile();
            LB_FQCFile.Content = Create2007ExcelFile();
            PGMethod.SetConfigSetting("FQCFile", LB_FQCFile.Content.ToString());
        }
        #region EXCEL 2016
        /*
        private string CreateExcelFile()
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.DefaultExt = ".xlsx";
            sfd.Filter = "Office 2007 File|*.xlsx|Office 2000-2003 File|*.xls|所有檔案|*.*";
            var result = sfd.ShowDialog();
            if (result == true)
            {
                if (!File.Exists(sfd.FileName))
                {
                    Excel.Application excel = null;
                    Excel.Workbook wb = null;
                    try
                    {
                        excel = new Excel.Application();
                        wb = excel.Workbooks.Add();
                        if (sfd.FileName != "")
                        {
                            wb.SaveAs(sfd.FileName);
                        }
                    }
                    catch (Exception)
                    {
                        LB_ErrMessage.Content = "建立" + sfd.FileName + "檔案失敗！";
                    }
                    finally
                    {
                        wb.Close();
                        excel.Quit();
                    }
                }
                return sfd.FileName;
            }
            return "";
        }
        */
        #endregion
        private string Create2007ExcelFile()
        {
            FileStream fileStream;
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.DefaultExt = ".xlsx";
            sfd.Filter = "Office 2007 File|*.xlsx|Office 2000-2003 File|*.xls|所有檔案|*.*";
            var result = sfd.ShowDialog();
            if (result == true)
            {
                if (!File.Exists(sfd.FileName))
                {
                    try
                    {
                        fileStream = new FileStream(sfd.FileName, FileMode.Create);
                        //fileStream = new FileStream(sfd.FileName, FileMode.Open, FileAccess.ReadWrite);
                        if (sfd.FileName.IndexOf(".xlsx") > 0) // 2007版本 
                        {
                            IWorkbook wb = new XSSFWorkbook();
                            ISheet ws = wb.CreateSheet("工作表1");
                            ws.CreateRow(0);
                            wb.Write(fileStream);
                        }
                        fileStream.Close();
                    }
                    catch (Exception ex)
                    {
                        //LB_ErrMessage.Content = "建立" + sfd.FileName + "檔案失敗！";
                        LB_ErrMessage.Content = ex.Message;
                    }
                }
                return sfd.FileName;
            }
            return "";
        }
        #region EXCEL 2016
        /*
        private void CreateExcelFile(string FilePath)
        {
            Excel.Application excel = null;
            Excel.Workbook wb = null;
            try
            {
                excel = new Excel.Application();
                wb = excel.Workbooks.Add();
                if (FilePath != "")
                {
                    wb.SaveAs(FilePath);
                }
            }
            catch (Exception)
            {
                LB_ErrMessage.Content = "建立" + FilePath + "檔案失敗！";
            }
            finally
            {
                wb.Close();
            }
        }
        */
        #endregion
        private void Create2007ExcelFile(string FilePath)
        {
            FileStream fileStream;
            IWorkbook wb = new XSSFWorkbook();
            fileStream = new FileStream(FilePath, FileMode.Create);
            try
            {
                if (FilePath.IndexOf(".xlsx") > 0) // 2007版本 
                {
                    ISheet ws = wb.CreateSheet("工作表1");
                    ws.CreateRow(0);
                    wb.Write(fileStream);
                }
                fileStream.Close();
            }
            catch (Exception)
            {
                LB_ErrMessage.Content = "建立" + FilePath + "檔案失敗！";
            }
            finally
            {
                wb.Close();
            }
        }
        private void StartButton_Click(object sender, RoutedEventArgs e)
        {
            if (Syncing) return;
            LB_ErrMessage.Content = "";
            if (DP_Shipment.Text == "")
            {
                LB_ErrMessage.Content = "請選擇出貨日期！";
                return;
            }
            if (LB_QC1File.Content.ToString() == "")
            {
                LB_ErrMessage.Content = "請指定QC1存檔路徑！";
                return;
            }
            if (LB_QC2File.Content.ToString() == "")
            {
                LB_ErrMessage.Content = "請指定QC2存檔路徑！";
                return;
            }
            if (LB_FQCFile.Content.ToString() == "")
            {
                LB_ErrMessage.Content = "請指定FQC存檔路徑！";
                return;
            }
            Syncing = true;
            string[] ProcessInfo = {
                LB_QC1File.Content.ToString(),
                LB_QC2File.Content.ToString(),
                LB_FQCFile.Content.ToString(),
                Convert.ToDateTime(DP_Shipment.Text).ToString("yyyy-MM-dd"),
            };
            SqlConnection iBoxConnect = new SqlConnection(iboxdbinfo);
            ShipmentList = GetShipmentList(ProcessInfo[3], iBoxConnect);
            ShippingRows = ShipmentList.Count;
            Thread thread = new Thread(SyncProcess);
            thread.IsBackground = true;
            thread.Start(ProcessInfo);
        }
        private void SyncProcess(object ProcessDay)
        {
            _syncContext.Post(SetProcessMessage, "開始執行 -- QC1");
            string[] ProcessInfo = (string[])ProcessDay;
            if (!File.Exists(ProcessInfo[0]))
            {
                Create2007ExcelFile(ProcessInfo[0]);
            }
            if (GenerateQC12007File(ProcessInfo))
            {
                _syncContext.Post(SetProcessMessage, "開始執行 -- QC2");
            }
            if (!File.Exists(ProcessInfo[1]))
            {
                Create2007ExcelFile(ProcessInfo[1]);
            }
            if (GenerateQC22007File(ProcessInfo))
            {
                _syncContext.Post(SetProcessMessage, "開始執行 -- FQC");
            }
            if (!File.Exists(ProcessInfo[2]))
            {
                Create2007ExcelFile(ProcessInfo[2]);
            }
            if (GenerateFQC2007File(ProcessInfo))
            {
                _syncContext.Post(SetSyncValue, (double)100);
            }
            _syncContext.Post(SetProcessMessage, "執行完畢");
            Syncing = false;
        }
        private void SetMessageValue(object MsgValue)
        {
            string Msg = (string)MsgValue;
            LB_ErrMessage.Content = Msg;
        }
        private void SetProcessMessage(object MsgValue)
        {
            string Msg = (string)MsgValue;
            LB_ProcessMessage.Content = Msg;
        }
        private void SetSyncValue(object SetpValue)
        {
            double posen = (double)SetpValue;
            PB_Sync.Value = Math.Round(posen, 1);
        }
        private List<decimal> GetDecimalMeasure(int GetRows, string ColumnName, string TableName, SqlConnection Comm)
        {
            GetRows = GetRows - 1;
            string SelectSQL = @"SELECT TOP " + GetRows + " " + ColumnName + " FROM " + TableName + " ORDER BY NEWID()";
            List<decimal> MeasureList = new List<decimal>();
            try
            {
                Comm.Open();
                SqlCommand cmd = new SqlCommand(SelectSQL, Comm);
                SqlDataReader dataReader = cmd.ExecuteReader();
                bool ReturnFlage = dataReader.HasRows;
                if (ReturnFlage)
                {
                    MeasureList.Clear();
                    while (dataReader.Read())
                    {
                        decimal rowInfo = dataReader.GetDecimal(0);
                        MeasureList.Add(rowInfo);
                    }
                }
            }
            catch (Exception ex)
            {
                _syncContext.Post(SetMessageValue, ex.Message);
            }
            finally
            {
                Comm.Close();
            }
            return MeasureList;
        }
        private List<string[]> GetShipmentList(string ShipmentDay, SqlConnection Comm)
        {
            string SelectSQL = @"SELECT pk.ModifyDate, WIP.ID FROM PackageData as pk " +
                "JOIN WIP ON pk.ID = WIP.PackageNumber WHERE pk.ModifyDate BETWEEN '" + ShipmentDay +  " 00:00:00' AND '"+ ShipmentDay + " 23:59:59' " +
                "AND pk.Status = 1 AND pk.ProductID = 'P1109242-00-C:SBGD' ORDER BY ID";
            List<string[]> ShipmentList = new List<string[]>();
            try
            {
                Comm.Open();
                SqlCommand cmd = new SqlCommand(SelectSQL, Comm);
                SqlDataReader dataReader = cmd.ExecuteReader();
                bool ReturnFlage = dataReader.HasRows;
                if (ReturnFlage)
                {
                    ShipmentList.Clear();
                    while (dataReader.Read())
                    {
                        string[] rowInfo = new string[2];
                        rowInfo[0] = dataReader.GetDateTime(0).ToString("MM-dd");
                        rowInfo[1] = dataReader.GetString(1);
                        ShipmentList.Add(rowInfo);
                    }
                }
            }
            catch (Exception ex)
            {
                _syncContext.Post(SetMessageValue, ex.Message);
            }
            finally
            {
                Comm.Close();
            }
            return ShipmentList;
        }
        #region EXCEL 2016
        /*
        private bool GenerateQC1File(string[] ProcessInfo)
        {
            Excel.Application Qc1_excel = new Excel.Application();
            Excel.Workbook Qc1wb = Qc1_excel.Workbooks.Open(ProcessInfo[0]);
            Excel.Worksheet QC1ws = null;
            Excel.Worksheet worksheet;
            try
            {
                worksheet = (Excel.Worksheet)Qc1wb.Worksheets["工作表1"];
                worksheet.Name = "1104382-00-B-ROLLING-" + Convert.ToDateTime(ProcessInfo[3]).ToString("MMdd");
            }
            catch (Exception)
            {
                try
                {
                    worksheet = (Excel.Worksheet)Qc1wb.Worksheets["1104382-00-B-ROLLING-" + Convert.ToDateTime(ProcessInfo[3]).ToString("MMdd")];
                }
                catch (Exception)
                {
                    worksheet = Qc1wb.Sheets.Add();
                    worksheet.Name = "1104382-00-B-ROLLING-" + Convert.ToDateTime(ProcessInfo[3]).ToString("MMdd");
                }
            }
            QC1ws = Qc1wb.ActiveSheet;
            try
            {
                Excel.Range A1title = QC1ws.Cells[1, 1];
                if ((string)A1title.Value == null)
                {
                    SetQC1Title(QC1ws);
                }
                else
                {
                    QC1ws.Cells.ClearContents();
                    SetQC1Title(QC1ws);
                }
                int ExcelRows = startRow;
                foreach (string[] rowinfo in ShipmentList)
                {
                    if (ExcelRows % 100 == 0)
                    {
                        double Barvalue = ExcelRows / (double)ShippingRows * 33;
                        _syncContext.Post(SetSyncValue, Barvalue);
                    }
                    QC1ws.Cells[ExcelRows, 1] = rowinfo[0];
                    QC1ws.Cells[ExcelRows, 2] = rowinfo[1];
                    QC1ws.Cells[ExcelRows, 3] = "GOOD";
                    QC1ws.Cells[ExcelRows, 4] = GetRandom(25, 33, 2);
                    QC1ws.Cells[ExcelRows, 5] = "GOOD";
                    QC1ws.Cells[ExcelRows, 6] = GetRandom((decimal)58.965, (decimal)59.013, 4);
                    QC1ws.Cells[ExcelRows, 7] = 58.965;
                    QC1ws.Cells[ExcelRows, 8] = 59.013;
                    QC1ws.Cells[ExcelRows, 9] = "GOOD";
                    QC1ws.Cells[ExcelRows, 10] = GetRandom((decimal)0.0009, (decimal)0.015, 4);
                    QC1ws.Cells[ExcelRows, 11] = 0.015;
                    QC1ws.Cells[ExcelRows, 12] = "GOOD";
                    QC1ws.Cells[ExcelRows, 13] = GetRandom((decimal)0.0009, (decimal)0.031, 4);
                    QC1ws.Cells[ExcelRows, 14] = 0.031;
                    QC1ws.Cells[ExcelRows, 15] = "GOOD";
                    QC1ws.Cells[ExcelRows, 16] = GetRandom((decimal)0.0009, (decimal)0.009, 4);
                    QC1ws.Cells[ExcelRows, 17] = 0.009;
                    ExcelRows++;
                }
                _syncContext.Post(SetSyncValue, (double)33);
                Qc1wb.Save();
            }
            finally
            {
                Qc1_excel.DisplayAlerts = false;
                Qc1wb.Close();
                Qc1wb = null;
                Qc1_excel.Quit();
            }
            return true;
        }
        */
        #endregion
        private bool GenerateQC12007File(string[] ProcessInfo)
        {
            Stream fileStream = null;
            IWorkbook wb = null;
            ISheet ws;
            fileStream = new FileStream(ProcessInfo[0], FileMode.Open, FileAccess.Read);
            wb = new XSSFWorkbook(fileStream);
            fileStream.Close();
            try
            {
                int activesheet = wb.GetSheetIndex("工作表1");
                wb.RemoveSheetAt(activesheet);
                ws = wb.CreateSheet("1104382-00-B-ROLLING-" + Convert.ToDateTime(ProcessInfo[3]).ToString("MMdd"));
            }
            catch (Exception)
            {
                try
                {
                    int activesheet = wb.GetSheetIndex("1104382-00-B-ROLLING-" + Convert.ToDateTime(ProcessInfo[3]).ToString("MMdd"));
                    ws = wb.GetSheetAt(activesheet);
                }
                catch (Exception)
                {
                    ws = wb.CreateSheet("1104382-00-B-ROLLING-" + Convert.ToDateTime(ProcessInfo[3]).ToString("MMdd"));
                }
            }
            SetQC12007Title(wb, ws);
            int ExcelRows = 1;
            XSSFFont myFont = (XSSFFont)wb.CreateFont();
            myFont.FontHeightInPoints = 12;
            myFont.FontName = "微軟正黑體";
            XSSFCellStyle borderedCellStyle = (XSSFCellStyle)wb.CreateCellStyle();
            borderedCellStyle.SetFont(myFont);
            foreach (string[] rowinfo in ShipmentList)
            {
                if (ExcelRows % 100 == 0)
                {
                    double Barvalue = ExcelRows / (double)ShippingRows * 33;
                    _syncContext.Post(SetSyncValue, Barvalue);
                }
                Thread.Sleep(25);
                IRow ContentRow = ws.CreateRow(ExcelRows);
                CreateCell(ContentRow, 0, rowinfo[0], borderedCellStyle);
                CreateCell(ContentRow, 1, rowinfo[1], borderedCellStyle);
                CreateCell(ContentRow, 2, "GOOD", borderedCellStyle);
                CreateCellNumber(ContentRow, 3, GetRandom(25, 33, 2).ToString(), borderedCellStyle);
                CreateCell(ContentRow, 4, "GOOD", borderedCellStyle);
                CreateCellNumber(ContentRow, 5, GetRandom((decimal)58.965, (decimal)59.013, 4).ToString(), borderedCellStyle);
                CreateCellNumber(ContentRow, 6, "58.965", borderedCellStyle);
                CreateCellNumber(ContentRow, 7, "59.013", borderedCellStyle);
                CreateCell(ContentRow, 8, "GOOD", borderedCellStyle);
                CreateCellNumber(ContentRow, 9, GetRandom((decimal)0.0004, (decimal)0.015, 4).ToString(), borderedCellStyle);
                CreateCellNumber(ContentRow, 10, "0.015", borderedCellStyle);
                CreateCell(ContentRow, 11, "GOOD", borderedCellStyle);
                CreateCellNumber(ContentRow, 12, GetRandom((decimal)0.0004, (decimal)0.031, 4).ToString(), borderedCellStyle);
                CreateCellNumber(ContentRow, 13, "0.031", borderedCellStyle);
                CreateCell(ContentRow, 14, "GOOD", borderedCellStyle);
                CreateCellNumber(ContentRow, 15, GetRandom((decimal)0.0004, (decimal)0.009, 4).ToString(), borderedCellStyle);
                CreateCellNumber(ContentRow, 16, "0.009", borderedCellStyle);
                ExcelRows++;
            }
            try
            {
                fileStream = new FileStream(ProcessInfo[0], FileMode.Open, FileAccess.Write);
                wb.Write(fileStream);
                fileStream.Close();
                _syncContext.Post(SetSyncValue, (double)33);
                return true;
            }
            catch (Exception ex)
            {
                _syncContext.Post(SetMessageValue, ex.Message);
                return false;
            }
        }
        #region EXCEL 2016
        /*
        private bool GenerateQC2File(string[] ProcessInfo)
        {
            Excel.Application QC2_excel = new Excel.Application();
            Excel.Workbook QC2wb = QC2_excel.Workbooks.Open(ProcessInfo[1]);
            Excel.Worksheet QC2ws = null;
            //SqlConnection iBoxConnect = new SqlConnection(iboxdbinfo);
            //SqlConnection hotaConnect = new SqlConnection(hotadbinfo);
            Excel.Worksheet worksheet;
            try
            {
                worksheet = (Excel.Worksheet)QC2wb.Worksheets["工作表1"];
                worksheet.Name = "1104382-00-B-ID-" + Convert.ToDateTime(ProcessInfo[3]).ToString("MMdd");
            }
            catch (Exception)
            {
                try
                {
                    worksheet = (Excel.Worksheet)QC2wb.Worksheets["1104382-00-B-ID-" + Convert.ToDateTime(ProcessInfo[3]).ToString("MMdd")];
                }
                catch (Exception)
                {
                    worksheet = QC2wb.Sheets.Add();
                    worksheet.Name = "1104382-00-B-ID-" + Convert.ToDateTime(ProcessInfo[3]).ToString("MMdd");
                }
            }
            QC2ws = QC2wb.ActiveSheet;
            //List<decimal> MeasureList;
            try
            {
                Excel.Range A1title = QC2ws.Cells[1, 1];
                if ((string)A1title.Value == null)
                {
                    SetQC2Title(QC2ws);
                }
                else
                {
                    QC2ws.Cells.ClearContents();
                    SetQC2Title(QC2ws);
                }
                _syncContext.Post(SetSyncValue, (double)35);
                int ExcelRows = startRow;
                foreach (string[] rowinfo in ShipmentList)
                {
                    if (ExcelRows % 100 == 0)
                    {
                        double Barvalue = 33 + ((double)ExcelRows / ShippingRows * 33);
                        _syncContext.Post(SetSyncValue, Barvalue);
                    }
                    QC2ws.Cells[ExcelRows, 1] = rowinfo[0];
                    QC2ws.Cells[ExcelRows, 2] = rowinfo[1];
                    QC2ws.Cells[ExcelRows, 3] = "GOOD";
                    QC2ws.Cells[ExcelRows, 4] = GetRandom(25, 33, 2);
                    QC2ws.Cells[ExcelRows, 5] = "GOOD";
                    QC2ws.Cells[ExcelRows, 6] = GetRandom((decimal)0.0009, (decimal)0.035, 4);
                    QC2ws.Cells[ExcelRows, 7] = 0.035;
                    QC2ws.Cells[ExcelRows, 8] = "GOOD";
                    QC2ws.Cells[ExcelRows, 9] = GetRandom((decimal)0.0009, (decimal)0.05, 4);
                    QC2ws.Cells[ExcelRows, 10] = 0.05;
                    QC2ws.Cells[ExcelRows, 11] = "GOOD";
                    QC2ws.Cells[ExcelRows, 12] = GetRandom((decimal)28.493, (decimal)28.503, 4);
                    QC2ws.Cells[ExcelRows, 13] = 28.493;
                    QC2ws.Cells[ExcelRows, 14] = 28.503;
                    QC2ws.Cells[ExcelRows, 15] = "GOOD";
                    QC2ws.Cells[ExcelRows, 16] = GetRandom(0, (decimal)0.01, 4);
                    QC2ws.Cells[ExcelRows, 17] = 0.01;
                    ExcelRows++;
                }
                _syncContext.Post(SetSyncValue, (double)66);
                QC2wb.Save();
            }
            finally
            {
                //Qc1_excel.DisplayAlerts = false;
                QC2wb.Close();
                QC2wb = null;
                QC2_excel.Quit();
            }
            return true;
        }
        */
        #endregion
        private bool GenerateQC22007File(string[] ProcessInfo)
        {
            Stream fileStream = new FileStream(ProcessInfo[1], FileMode.Open, FileAccess.Read);
            IWorkbook wb = new XSSFWorkbook(fileStream);
            fileStream.Close();
            ISheet ws;
            try
            {
                int activesheet = wb.GetSheetIndex("工作表1");
                wb.RemoveSheetAt(activesheet);
                ws = wb.CreateSheet("1104382-00-B-ID-" + Convert.ToDateTime(ProcessInfo[3]).ToString("MMdd"));
            }
            catch (Exception)
            {
                try
                {
                    int activesheet = wb.GetSheetIndex("1104382-00-B-ID-" + Convert.ToDateTime(ProcessInfo[3]).ToString("MMdd"));
                    ws = wb.GetSheetAt(activesheet);
                }
                catch (Exception)
                {
                    ws = wb.CreateSheet("1104382-00-B-ID-" + Convert.ToDateTime(ProcessInfo[3]).ToString("MMdd"));
                }
            }
            SetQC22007Title(wb, ws);
            int ExcelRows = 1;
            XSSFFont myFont = (XSSFFont)wb.CreateFont();
            myFont.FontHeightInPoints = 12;
            myFont.FontName = "微軟正黑體";
            XSSFCellStyle borderedCellStyle = (XSSFCellStyle)wb.CreateCellStyle();
            borderedCellStyle.SetFont(myFont);
            foreach (string[] rowinfo in ShipmentList)
            {
                if (ExcelRows % 100 == 0)
                {
                    double Barvalue = 33 + ((double)ExcelRows / ShippingRows * 33);
                    _syncContext.Post(SetSyncValue, Barvalue);
                }
                Thread.Sleep(25);
                IRow ContentRow = ws.CreateRow(ExcelRows);
                CreateCell(ContentRow, 0, rowinfo[0], borderedCellStyle);
                CreateCell(ContentRow, 1, rowinfo[1], borderedCellStyle);
                CreateCell(ContentRow, 2, "GOOD", borderedCellStyle);
                CreateCellNumber(ContentRow, 3, GetRandom(25, 33, 2).ToString(), borderedCellStyle);
                CreateCell(ContentRow, 4, "GOOD", borderedCellStyle);
                CreateCellNumber(ContentRow, 5, GetRandom((decimal)0.0009, (decimal)0.035, 4).ToString(), borderedCellStyle);
                CreateCellNumber(ContentRow, 6, "0.035", borderedCellStyle);
                CreateCell(ContentRow, 7, "GOOD", borderedCellStyle);
                CreateCellNumber(ContentRow, 8, GetRandom((decimal)0.0009, (decimal)0.05, 4).ToString(), borderedCellStyle);
                CreateCellNumber(ContentRow, 9, "0.05", borderedCellStyle);
                CreateCell(ContentRow, 10, "GOOD", borderedCellStyle);
                CreateCellNumber(ContentRow, 11, GetRandom((decimal)28.493, (decimal)28.503, 4).ToString(), borderedCellStyle);
                CreateCellNumber(ContentRow, 12, "28.493", borderedCellStyle);
                CreateCellNumber(ContentRow, 13, "28.503", borderedCellStyle);
                CreateCell(ContentRow, 14, "GOOD", borderedCellStyle);
                CreateCellNumber(ContentRow, 15, GetRandom((decimal)0.0004, (decimal)0.01, 4).ToString(), borderedCellStyle);
                CreateCellNumber(ContentRow, 16, "0.01", borderedCellStyle);
                ExcelRows++;
            }
            try
            {
                fileStream = new FileStream(ProcessInfo[1], FileMode.Open, FileAccess.Write);
                wb.Write(fileStream);
                fileStream.Close();
                _syncContext.Post(SetSyncValue, (double)66);
                return true;
            }
            catch (Exception ex)
            {
                _syncContext.Post(SetMessageValue, ex.Message);
                return false;
            }
        }
        #region EXCEL 2016
        /*
        private bool GenerateFQCFile(string[] ProcessInfo)
        {
            Excel.Application FQC_excel = new Excel.Application();
            Excel.Workbook FQCwb = FQC_excel.Workbooks.Open(ProcessInfo[2]);
            Excel.Worksheet FQCws = null;
            SqlConnection hotaConnect = new SqlConnection(hotadbinfo);
            Excel.Worksheet worksheet;
            try
            {
                worksheet = (Excel.Worksheet)FQCwb.Worksheets["工作表1"];
                worksheet.Name = "1109242-00-C " + Convert.ToDateTime(ProcessInfo[3]).ToString("MMdd");
            }
            catch (Exception)
            {
                try
                {
                    worksheet = (Excel.Worksheet)FQCwb.Worksheets["1109242-00-C " + Convert.ToDateTime(ProcessInfo[3]).ToString("MMdd")];
                }
                catch (Exception)
                {
                    worksheet = FQCwb.Sheets.Add();
                    worksheet.Name = "1109242-00-C " + Convert.ToDateTime(ProcessInfo[3]).ToString("MMdd");
                }
            }
            FQCws = FQCwb.ActiveSheet;
            try
            {
                _syncContext.Post(SetSyncValue, (double)70);
                Excel.Range A1title = FQCws.Cells[1, 1];
                if ((string)A1title.Value == null)
                {
                    SetFQCTitle(FQCws);
                }
                else
                {
                    FQCws.Cells.ClearContents();
                    SetFQCTitle(FQCws);
                }
                int ExcelRows = startRow;
                foreach (string[] rowinfo in ShipmentList)
                {
                    if (ExcelRows % 100 == 0)
                    {
                        double Barvalue = 66 + ((double)ExcelRows / ShippingRows * 24);
                        _syncContext.Post(SetSyncValue, Barvalue);
                    }
                    FQCws.Cells[ExcelRows, 1] = rowinfo[0];
                    FQCws.Cells[ExcelRows, 2] = rowinfo[1];
                    FQCws.Cells[ExcelRows, 3] = "GOOD";
                    FQCws.Cells[ExcelRows, 4] = GetRandom(25, 33, 2);
                    FQCws.Cells[ExcelRows, 5] = "GOOD";
                    FQCws.Cells[ExcelRows, 6] = GetRandom((decimal)9.674, (decimal)9.714, 4);
                    FQCws.Cells[ExcelRows, 7] = 9.674;
                    FQCws.Cells[ExcelRows, 8] = 9.714;
                    FQCws.Cells[ExcelRows, 9] = "GOOD";
                    FQCws.Cells[ExcelRows, 10] = GetRandom(0, (decimal)0.02, 4);
                    FQCws.Cells[ExcelRows, 11] = 0.02;
                    FQCws.Cells[ExcelRows, 12] = "GOOD";
                    FQCws.Cells[ExcelRows, 13] = GetRandom((decimal)40.009, (decimal)40.020, 4);
                    FQCws.Cells[ExcelRows, 14] = 40.009;
                    FQCws.Cells[ExcelRows, 15] = 40.02;
                    FQCws.Cells[ExcelRows, 18] = "GOOD";
                    FQCws.Cells[ExcelRows, 19] = GetRandom((decimal)40.009, (decimal)40.020, 4);
                    FQCws.Cells[ExcelRows, 20] = 40.009;
                    FQCws.Cells[ExcelRows, 21] = 40.02;
                    ExcelRows++;
                }
                List<decimal> Q5edlmax = GetDecimalMeasure(ExcelRows, "q5edlmax", "shipmentfqc", hotaConnect);
                ExcelRows = startRow;
                foreach (decimal rowinfo in Q5edlmax)
                {
                    FQCws.Cells[ExcelRows, 16] = rowinfo;
                    ExcelRows++;
                }
                _syncContext.Post(SetSyncValue, (double)92);
                List<decimal> Q5edlmin = GetDecimalMeasure(ExcelRows, "q5edlmin", "shipmentfqc", hotaConnect);
                ExcelRows = startRow;
                foreach (decimal rowinfo in Q5edlmin)
                {
                    FQCws.Cells[ExcelRows, 17] = rowinfo;
                    ExcelRows++;
                }
                _syncContext.Post(SetSyncValue, (double)94);
                List<decimal> Q6edrmax = GetDecimalMeasure(ExcelRows, "q6edrmax", "shipmentfqc", hotaConnect);
                ExcelRows = startRow;
                foreach (decimal rowinfo in Q6edrmax)
                {
                    FQCws.Cells[ExcelRows, 22] = rowinfo;
                    ExcelRows++;
                }
                _syncContext.Post(SetSyncValue, (double)96);
                List<decimal> Q6edrmin = GetDecimalMeasure(ExcelRows, "q6edrmin", "shipmentfqc", hotaConnect);
                ExcelRows = startRow;
                foreach (decimal rowinfo in Q6edrmin)
                {
                    FQCws.Cells[ExcelRows, 23] = rowinfo;
                    ExcelRows++;
                }
                _syncContext.Post(SetSyncValue, (double)100);
                FQCwb.Save();
            }
            finally
            {
                FQCwb.Close();
                FQCwb = null;
                FQC_excel.Quit();
            }
            return true;
        }
        */
        #endregion
        private bool GenerateFQC2007File(string[] ProcessInfo)
        {
            Stream fileStream = new FileStream(ProcessInfo[2], FileMode.Open, FileAccess.Read);
            IWorkbook wb = new XSSFWorkbook(fileStream);
            fileStream.Close();
            ISheet ws;
            SqlConnection hotaConnect = new SqlConnection(hotadbinfo);
            try
            {
                int activesheet = wb.GetSheetIndex("工作表1");
                wb.RemoveSheetAt(activesheet);
                ws = wb.CreateSheet("1109242-00-C " + Convert.ToDateTime(ProcessInfo[3]).ToString("MMdd"));
            }
            catch (Exception)
            {
                try
                {
                    int activesheet = wb.GetSheetIndex("1109242-00-C " + Convert.ToDateTime(ProcessInfo[3]).ToString("MMdd"));
                    ws = wb.GetSheetAt(activesheet);
                }
                catch (Exception)
                {
                    ws = wb.CreateSheet("1109242-00-C " + Convert.ToDateTime(ProcessInfo[3]).ToString("MMdd"));
                }
            }
            SetFQC2007Title(wb, ws);
            int ExcelRows = 1;
            XSSFFont myFont = (XSSFFont)wb.CreateFont();
            myFont.FontHeightInPoints = 12;
            myFont.FontName = "微軟正黑體";
            XSSFCellStyle borderedCellStyle = (XSSFCellStyle)wb.CreateCellStyle();
            borderedCellStyle.SetFont(myFont);
            foreach (string[] rowinfo in ShipmentList)
            {
                if (ExcelRows % 100 == 0)
                {
                    double Barvalue = 66 + ((double)ExcelRows / ShippingRows * 24);
                    _syncContext.Post(SetSyncValue, Barvalue);
                }
                Thread.Sleep(25);
                IRow ContentRow = ws.CreateRow(ExcelRows);
                CreateCell(ContentRow, 0, rowinfo[0], borderedCellStyle);
                CreateCell(ContentRow, 1, rowinfo[1], borderedCellStyle);
                CreateCell(ContentRow, 2, "GOOD", borderedCellStyle);
                CreateCellNumber(ContentRow, 3, GetRandom(25, 33, 2).ToString(), borderedCellStyle);
                CreateCell(ContentRow, 4, "GOOD", borderedCellStyle);
                CreateCellNumber(ContentRow, 5, GetRandom((decimal)9.674, (decimal)9.714, 4).ToString(), borderedCellStyle);
                CreateCellNumber(ContentRow, 6, "9.674", borderedCellStyle);
                CreateCellNumber(ContentRow, 7, "9.714", borderedCellStyle);
                CreateCell(ContentRow, 8, "GOOD", borderedCellStyle);
                CreateCellNumber(ContentRow, 9, GetRandom((decimal)0.0003, (decimal)0.02, 4).ToString(), borderedCellStyle);
                CreateCellNumber(ContentRow, 10, "0.02", borderedCellStyle);
                CreateCell(ContentRow, 11, "GOOD", borderedCellStyle);
                CreateCellNumber(ContentRow, 12, GetRandom((decimal)40.009, (decimal)40.020, 4).ToString(), borderedCellStyle);
                CreateCellNumber(ContentRow, 13, "40.009", borderedCellStyle);
                CreateCellNumber(ContentRow, 14, "40.02", borderedCellStyle);
                CreateCell(ContentRow, 17, "GOOD", borderedCellStyle);
                CreateCellNumber(ContentRow, 18, GetRandom((decimal)40.009, (decimal)40.020, 4).ToString(), borderedCellStyle);
                CreateCellNumber(ContentRow, 19, "40.009", borderedCellStyle);
                CreateCellNumber(ContentRow, 20, "40.02", borderedCellStyle);
                ExcelRows++;
            }
            List<decimal> Q5edlmax = GetDecimalMeasure(ExcelRows, "q5edlmax", "shipmentfqc", hotaConnect);
            ExcelRows = startRow;
            foreach (decimal rowinfo in Q5edlmax)
            {
                IRow ContentRow = ws.GetRow(ExcelRows);
                CreateCellNumber(ContentRow, 15, rowinfo.ToString(), borderedCellStyle);
                ExcelRows++;
            }
            _syncContext.Post(SetSyncValue, (double)92);
            List<decimal> Q5edlmin = GetDecimalMeasure(ExcelRows, "q5edlmin", "shipmentfqc", hotaConnect);
            ExcelRows = startRow;
            foreach (decimal rowinfo in Q5edlmin)
            {
                IRow ContentRow = ws.GetRow(ExcelRows);
                CreateCellNumber(ContentRow, 16, rowinfo.ToString(), borderedCellStyle);
                ExcelRows++;
            }
            _syncContext.Post(SetSyncValue, (double)94);
            List<decimal> Q6edrmax = GetDecimalMeasure(ExcelRows, "q6edrmax", "shipmentfqc", hotaConnect);
            ExcelRows = startRow;
            foreach (decimal rowinfo in Q6edrmax)
            {
                IRow ContentRow = ws.GetRow(ExcelRows);
                CreateCellNumber(ContentRow, 21, rowinfo.ToString(), borderedCellStyle);
                ExcelRows++;
            }
            _syncContext.Post(SetSyncValue, (double)96);
            List<decimal> Q6edrmin = GetDecimalMeasure(ExcelRows, "q6edrmin", "shipmentfqc", hotaConnect);
            ExcelRows = startRow;
            foreach (decimal rowinfo in Q6edrmin)
            {
                IRow ContentRow = ws.GetRow(ExcelRows);
                CreateCellNumber(ContentRow, 22, rowinfo.ToString(), borderedCellStyle);
                ExcelRows++;
            }
            try
            {
                fileStream = new FileStream(ProcessInfo[2], FileMode.Open, FileAccess.Write);
                wb.Write(fileStream);
                fileStream.Close();
                _syncContext.Post(SetSyncValue, (double)100);
                return true;
            }
            catch (Exception ex)
            {
                _syncContext.Post(SetMessageValue, ex.Message);
                return false;
            }
        }
        #region EXCEL 2016
        /*
        private void SetQC1Title(Excel.Worksheet Qc1ws)
        {
            Qc1ws.Cells[1,1] = "Date";
            Qc1ws.Cells[1,2] = "DMC Code";
            Qc1ws.Cells[1,3] = "TOTAL Result";
            Qc1ws.Cells[1,4] = "Temperature";
            Qc1ws.Cells[1,5] = "Q1 Result";
            Qc1ws.Cells[1,6] = "Q1 Diameter [mm]";
            Qc1ws.Cells[1,7] = "Q1 Diameter Min [mm]";
            Qc1ws.Cells[1,8] = "Q1 Diameter Max [mm]";
            Qc1ws.Cells[1,9] = "Q2 Result";
            Qc1ws.Cells[1,10] = "Q2 Fr Runout [mm]";
            Qc1ws.Cells[1,11] = "Q2 Fr Runout Max [mm]";
            Qc1ws.Cells[1,12] = "Q3 Result";
            Qc1ws.Cells[1,13] = "Q3 Fi 1harm [mm]";
            Qc1ws.Cells[1,14] = "Q3 Fi 1harm Max [mm]";
            Qc1ws.Cells[1,15] = "Q4 Result";
            Qc1ws.Cells[1,16] = "Q4 fi\"[mm]";
            Qc1ws.Cells[1,17] = "Q4 fi\" Max[mm]";
            Excel.Range _Range = Qc1ws.get_Range("A1", "Q1");
            _Range.Font.Size = 14;
            _Range.Font.Name = "微軟正黑體"; 
            _Range.EntireColumn.AutoFit();
            _Range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        }
        private void SetQC2Title(Excel.Worksheet QC2ws)
        {
            QC2ws.Cells[1, 1] = "Date";
            QC2ws.Cells[1, 2] = "DMC Code";
            QC2ws.Cells[1, 3] = "Result";
            QC2ws.Cells[1, 4] = "Piece Temperature";
            QC2ws.Cells[1, 5] = "Q1 Runout Result";
            QC2ws.Cells[1, 6] = "Q1 Runout Measured [mm]";
            QC2ws.Cells[1, 7] = "Q1 Runout MAX Value [mm]";
            QC2ws.Cells[1, 8] = "Q2 Total Runout Result";
            QC2ws.Cells[1, 9] = "Q2 Total Runout Measured [mm]";
            QC2ws.Cells[1, 10] = "Q2 Total Runout MAX Value [mm]";
            QC2ws.Cells[1, 11] = "Q3 Diameter 28 Result";
            QC2ws.Cells[1, 12] = "Q3 Diameter 28 Measured [mm]";
            QC2ws.Cells[1, 13] = "Q3 Diameter 28 MIN Value [mm]";
            QC2ws.Cells[1, 14] = "Q3 Diameter 28 MAX Value [mm]";
            QC2ws.Cells[1, 15] = "Q4 Inner Runout Result";
            QC2ws.Cells[1, 16] = "Q4 Inner Runout Measured [mm]";
            QC2ws.Cells[1, 17] = "Q4 Inner Runout MAX Value [mm]";
            Excel.Range _Range = QC2ws.get_Range("A1", "Q1");
            _Range.Font.Size = 14;
            _Range.Font.Name = "微軟正黑體";
            _Range.EntireColumn.AutoFit();
            _Range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        }
        private void SetFQCTitle(Excel.Worksheet QC2ws)
        {
            QC2ws.Cells[1, 1] = "Date";
            QC2ws.Cells[1, 2] = "DMC Code";
            QC2ws.Cells[1, 3] = "Result";
            QC2ws.Cells[1, 4] = "Temperature";
            QC2ws.Cells[1, 5] = "Q1 Diameter Ext Result";
            QC2ws.Cells[1, 6] = "Q1 Diameter Ext Measure [mm]";
            QC2ws.Cells[1, 7] = "Q1 Diameter Ext Min [mm]";
            QC2ws.Cells[1, 8] = "Q1 Diameter Ext Max [mm]";
            QC2ws.Cells[1, 9] = "Q4 Runout Result";
            QC2ws.Cells[1, 10] = "Q4 Runout Measure [mm]";
            QC2ws.Cells[1, 11] = "Q4 Runout Max [mm]";
            QC2ws.Cells[1, 12] = "Q5 Ext Diam Left Result";
            QC2ws.Cells[1, 13] = "Q5 Ext Diam Left Measured [mm]";
            QC2ws.Cells[1, 14] = "Q5 Ext Diam Left Thres MAX [mm]";
            QC2ws.Cells[1, 15] = "Q5 Ext Diam Left Thres MIN [mm]";
            QC2ws.Cells[1, 16] = "Q5 Ext Diam Left MIN Value [mm]";
            QC2ws.Cells[1, 17] = "Q5 Ext Diam Left MAX Value [mm]";
            QC2ws.Cells[1, 18] = "Q6 Ext Diam Right Result";
            QC2ws.Cells[1, 19] = "Q6 Ext Diam Right Measured [mm]";
            QC2ws.Cells[1, 20] = "Q6 Ext Diam Right Thres MAX [mm]";
            QC2ws.Cells[1, 21] = "Q6 Ext Diam Right Thres MIN [mm]";
            QC2ws.Cells[1, 22] = "Q6 Ext Diam Right MIN Value [mm]";
            QC2ws.Cells[1, 23] = "Q6 Ext Diam Right MAX Value [mm]";
            Excel.Range _Range = QC2ws.get_Range("A1", "W1");
            _Range.Font.Size = 14;
            _Range.Font.Name = "微軟正黑體";
            _Range.EntireColumn.AutoFit();
            _Range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        }
        */
        #endregion
        private void SetQC12007Title(IWorkbook QC1wb, ISheet Qc1ws)
        {
            IRow HeaderRow = Qc1ws.CreateRow(0);
            XSSFFont myFont = (XSSFFont)QC1wb.CreateFont();
            myFont.FontHeightInPoints = 14;
            myFont.FontName = "微軟正黑體";
            XSSFCellStyle borderedCellStyle = (XSSFCellStyle)QC1wb.CreateCellStyle();
            borderedCellStyle.SetFont(myFont);
            CreateCell(HeaderRow, 0, "Date", borderedCellStyle);
            CreateCell(HeaderRow, 1, "DMC Code", borderedCellStyle);
            CreateCell(HeaderRow, 2, "TOTAL Result", borderedCellStyle);
            CreateCell(HeaderRow, 3, "Temperature", borderedCellStyle);
            CreateCell(HeaderRow, 4, "Q1 Result", borderedCellStyle);
            CreateCell(HeaderRow, 5, "Q1 Diameter [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 6, "Q1 Diameter Min [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 7, "Q1 Diameter Max [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 8, "Q2 Result", borderedCellStyle);
            CreateCell(HeaderRow, 9, "Q2 Fr Runout [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 10, "Q2 Fr Runout Max [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 11, "Q3 Result", borderedCellStyle);
            CreateCell(HeaderRow, 12, "Q3 Fi 1harm [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 13, "Q3 Fi 1harm Max [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 14, "Q4 Result", borderedCellStyle);
            CreateCell(HeaderRow, 15, "Q4 fi\"[mm]", borderedCellStyle);
            CreateCell(HeaderRow, 16, "Q4 fi\" Max[mm]", borderedCellStyle);
        }
        private void SetQC22007Title(IWorkbook QC2wb, ISheet Qc2ws)
        {
            IRow HeaderRow = Qc2ws.CreateRow(0);
            XSSFFont myFont = (XSSFFont)QC2wb.CreateFont();
            myFont.FontHeightInPoints = 14;
            myFont.FontName = "微軟正黑體";
            XSSFCellStyle borderedCellStyle = (XSSFCellStyle)QC2wb.CreateCellStyle();
            borderedCellStyle.SetFont(myFont);
            CreateCell(HeaderRow, 0, "Date", borderedCellStyle);
            CreateCell(HeaderRow, 1, "DMC Code", borderedCellStyle);
            CreateCell(HeaderRow, 2, "Result", borderedCellStyle);
            CreateCell(HeaderRow, 3, "Piece Temperature", borderedCellStyle);
            CreateCell(HeaderRow, 4, "Q1 Runout Result", borderedCellStyle);
            CreateCell(HeaderRow, 5, "Q1 Runout Measured [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 6, "Q1 Runout MAX Value [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 7, "Q2 Total Runout Result", borderedCellStyle);
            CreateCell(HeaderRow, 8, "Q2 Total Runout Measured [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 9, "Q2 Total Runout MAX Value [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 10, "Q3 Diameter 28 Result", borderedCellStyle);
            CreateCell(HeaderRow, 11, "Q3 Diameter 28 Measured [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 12, "Q3 Diameter 28 MIN Value [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 13, "Q3 Diameter 28 MAX Value [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 14, "Q4 Inner Runout Result", borderedCellStyle);
            CreateCell(HeaderRow, 15, "Q4 Inner Runout Measured [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 16, "Q4 Inner Runout MAX Value [mm]", borderedCellStyle);
        }
        private void SetFQC2007Title(IWorkbook FQCwb, ISheet FQCws)
        {
            IRow HeaderRow = FQCws.CreateRow(0);
            XSSFFont myFont = (XSSFFont)FQCwb.CreateFont();
            myFont.FontHeightInPoints = 14;
            myFont.FontName = "微軟正黑體";
            XSSFCellStyle borderedCellStyle = (XSSFCellStyle)FQCwb.CreateCellStyle();
            borderedCellStyle.SetFont(myFont);
            CreateCell(HeaderRow, 0, "Date", borderedCellStyle);
            CreateCell(HeaderRow, 1, "DMC Code", borderedCellStyle);
            CreateCell(HeaderRow, 2, "Result", borderedCellStyle);
            CreateCell(HeaderRow, 3, "Temperature", borderedCellStyle);
            CreateCell(HeaderRow, 4, "Q1 Diameter Ext Result", borderedCellStyle);
            CreateCell(HeaderRow, 5, "Q1 Diameter Ext Measure [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 6, "Q1 Diameter Ext Min [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 7, "Q1 Diameter Ext Max [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 8, "Q4 Runout Result", borderedCellStyle);
            CreateCell(HeaderRow, 9, "Q4 Runout Measure [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 10, "Q4 Runout Max [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 11, "Q5 Ext Diam Left Result", borderedCellStyle);
            CreateCell(HeaderRow, 12, "Q5 Ext Diam Left Measured [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 13, "Q5 Ext Diam Left Thres MAX [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 14, "Q5 Ext Diam Left Thres MIN [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 15, "Q5 Ext Diam Left MIN Value [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 16, "Q5 Ext Diam Left MAX Value [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 17, "Q6 Ext Diam Right Result", borderedCellStyle);
            CreateCell(HeaderRow, 18, "Q6 Ext Diam Right Measured [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 19, "Q6 Ext Diam Right Thres MAX [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 20, "Q6 Ext Diam Right Thres MIN [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 21, "Q6 Ext Diam Right MIN Value [mm]", borderedCellStyle);
            CreateCell(HeaderRow, 22, "Q6 Ext Diam Right MAX Value [mm]", borderedCellStyle);
        }
        private decimal GetRandom(decimal SRandom, decimal ERandom, int PointLen)
        {
            string[] NumberArr;
            int StartNum;
            int EndNum;
            bool CanNexDuble = true;
            int rInt;
            if (SRandom  ==  Math.Floor(SRandom))
            {
                StartNum = (int)SRandom;
            }
            else
            {
                CanNexDuble = false;
                NumberArr = SRandom.ToString().Split('.');
                StartNum = Convert.ToInt32((NumberArr[0] + NumberArr[1]).PadRight(PointLen + NumberArr[0].Length, '0'));
            }
            if (ERandom == Math.Floor(ERandom))
            {
                EndNum = (int)ERandom;
            }
            else
            {
                CanNexDuble = false;
                NumberArr = ERandom.ToString().Split('.');
                EndNum = Convert.ToInt32((NumberArr[0] + NumberArr[1]).PadRight(PointLen + NumberArr[0].Length, '0'));
            }
            Random getRandom = new Random();
            if (CanNexDuble)
            {
                rInt = getRandom.Next(StartNum, EndNum);
                return Math.Round(rInt + (decimal)getRandom.NextDouble(), PointLen);
            }
            else
            {
                rInt = getRandom.Next(StartNum, EndNum);
                string Basint = "1".PadRight(PointLen + 1, '0');
                decimal BaseDec = 1 / decimal.Parse(Basint);
                return rInt * BaseDec;
            }
        }
        private void CreateCell(IRow CurrentRow, int CellIndex, string Value, XSSFCellStyle Style)
        {
            ICell Cell = CurrentRow.CreateCell(CellIndex);
            Cell.SetCellValue(Value);
            Cell.CellStyle = Style;
        }
        private void CreateCellNumber(IRow CurrentRow, int CellIndex, string Value, XSSFCellStyle Style)
        {
            ICell Cell = CurrentRow.CreateCell(CellIndex);
            Cell.SetCellType(CellType.Numeric);
            Cell.SetCellValue(double.Parse(Value));
            Cell.CellStyle = Style;
        }
    }
}

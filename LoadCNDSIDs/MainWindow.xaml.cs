using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using System.Windows.Threading;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;
using System.Diagnostics;

namespace LoadCNDSIDs
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string fileName;
        string connectionString = ConfigurationManager.ConnectionStrings["CoastalCareConnectionString"].ToString();
        string[] allowedMCOs = new string[] { "Coastal Care", "ECBH" };
        public MainWindow()
        {
            InitializeComponent();
        }
        private void selectFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            bool? good = dialog.ShowDialog();
            if (good.HasValue && good.Value)
            {
                fileName = dialog.FileName;
            }
        }

        private void LoadFile_Click(object sender, RoutedEventArgs e)
        {
            Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
            {
                LoadData();
            }));

        }

        private void LoadData()
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Open(fileName);
            Microsoft.Office.Interop.Excel.Worksheet activeSheet = workbook.ActiveSheet;
            int lastRow = activeSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Row;
            int lastColumn = activeSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Column;

            List<string> mcos = new List<string>();
            mcos.AddRange(allowedMCOs);
            int last = 0;
            int first = 0;
            int middleCount = 0;
            int ssnCount = 0;
            int dobCount = 0;
            int genderCount = 0;
            int recordChanged = 0;
            int recordAdded = 0;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                DataTable tableCNDS = new DataTable();
                using (SqlDataAdapter adapter = new SqlDataAdapter("select * from CNDSIDs", connection))
                {
                    adapter.FillSchema(tableCNDS, SchemaType.Source);
                    adapter.Fill(tableCNDS);
                }
                DataTable tableLocal = new DataTable();
                using (SqlDataAdapter adapter = new SqlDataAdapter("select * from CNDSIDLocalIDs", connection))
                {
                    adapter.FillSchema(tableLocal, SchemaType.Source);
                    adapter.Fill(tableLocal);
                }
                //index starts from 1, first row is the header
                for (int index = 2; index <= lastRow; index++)
                {
                    Array values = (Array)activeSheet.get_Range("A" + index.ToString(), "J" + index.ToString()).Cells.Value;
                    string MCO = Convert.ToString(values.GetValue(1, 2));
                    string CNDSID = Convert.ToString(values.GetValue(1, 3));
                    string clientId = Convert.ToString(values.GetValue(1, 4));
                    int mcoCode = Convert.ToInt32(clientId.Substring(0, 5));
                    int client;
                    if (!int.TryParse(clientId.Substring(5), out client))
                        continue;

                    string lastName = Convert.ToString(values.GetValue(1, 5));
                    string firstName = Convert.ToString(values.GetValue(1, 6));
                    string middle = Convert.ToString(values.GetValue(1, 7));
                    DateTime dob = Convert.ToDateTime(values.GetValue(1, 8));
                    int ssn = Convert.ToInt32(values.GetValue(1, 9));
                    string gender = Convert.ToString(values.GetValue(1, 10));

                    DataRow row = tableCNDS.Rows.Find(CNDSID);
                    if (row == null)
                    {
                        row = tableCNDS.NewRow();
                        row["CNDSID"] = CNDSID;
                        row["LastName"] = lastName;
                        row["FirstName"] = firstName;
                        row["MiddleName"] = middle;
                        row["DOB"] = dob;
                        row["SSN"] = ssn;
                        row["Gender"] = gender;
                        tableCNDS.Rows.Add(row);

                        recordAdded++;
                    }
                    else
                    {
                        bool added = false;
                        if (!((string)row["LastName"]).Equals(lastName, StringComparison.InvariantCultureIgnoreCase))
                        {
                            row["LastName"] = lastName;
                            added = true;
                            recordChanged++;
                            last++;
                        }
                        if (!((string)row["FirstName"]).Equals(firstName, StringComparison.InvariantCultureIgnoreCase))
                        {
                            row["FirstName"] = firstName;
                            if (!added)
                            {
                                added = true;
                                recordChanged++;
                            }
                            first++;
                        }
                        if (!((string)row["MiddleName"]).Equals(middle, StringComparison.InvariantCultureIgnoreCase))
                        {
                            row["MiddleName"] = middle;
                            if (!added)
                            {
                                added = true;
                                recordChanged++;
                            }
                            middleCount++;
                        }
                        if (((DateTime)row["dob"]) != dob)
                        {
                            row["DOB"] = dob;
                            if (!added)
                            {
                                added = true;
                                recordChanged++;
                            }
                            dobCount++;
                        }
                        if (((int)row["ssn"]) != ssn)
                        {
                            row["SSN"] = ssn;
                            if (!added)
                            {
                                added = true;
                                recordChanged++;
                            }
                            ssnCount++;
                        }
                        if (!((string)row["Gender"]).Trim().Equals(gender, StringComparison.InvariantCultureIgnoreCase))
                        {
                            row["Gender"] = gender;
                            if (!added)
                            {
                                added = true;
                                recordChanged++;
                            }
                            genderCount++;
                        }
                    }
                    if (!mcos.Contains(MCO))
                        continue;

                    DataRow localRow = tableLocal.Rows.Find(new object[] { CNDSID, MCO });
                    if (localRow == null)
                    {
                        localRow = tableLocal.NewRow();
                        localRow["CNDSID"] = CNDSID;
                        localRow["MCO"] = MCO;
                        localRow["MCOCode"] = mcoCode;
                        localRow["ClientID"] = client;

                        tableLocal.Rows.Add(localRow);
                    }
                    else
                    {
                        localRow["MCOCode"] = mcoCode;
                        localRow["ClientID"] = client;
                    }
                }
                
                SqlTransaction transaction = connection.BeginTransaction();
                try
                {
                    using (SqlDataAdapter adapter = new SqlDataAdapter("select * from CNDSIDs", connection))
                    {
                        adapter.SelectCommand.Transaction = transaction;
                        SqlCommandBuilder builder = new SqlCommandBuilder(adapter);
                        adapter.Update(tableCNDS);
                    }
                    using (SqlDataAdapter adapter = new SqlDataAdapter("select * from CNDSIDLocalIDs", connection))
                    {
                        adapter.SelectCommand.Transaction = transaction;
                        SqlCommandBuilder builder = new SqlCommandBuilder(adapter);
                        adapter.Update(tableLocal);
                    }
                    transaction.Commit();
                }
                catch (Exception e)
                {
                    transaction.Rollback();
                    MessageBox.Show(e.ToString());
                }
                MessageBox.Show(string.Format("total records changed: {0} total record added: {1} \nChanged Last Name: {2}  First Name: {3}  Middle Name: {4}  DOB: {5}  SSN: {6} Gender: {7}", recordChanged, recordAdded, last, first, middleCount, dobCount, ssnCount, genderCount));
            }
        }
    }
}

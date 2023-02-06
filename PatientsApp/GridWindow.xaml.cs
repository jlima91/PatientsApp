using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace PatientsApp
{
    /// <summary>
    /// Interaction logic for GridWindow.xaml
    /// </summary>
    public partial class GridWindow : System.Windows.Window
    { 
        readonly SqlConnection sqlConnection;
        public GridWindow()
        {
            InitializeComponent();
            string connectionString = ConfigurationManager.ConnectionStrings["PatientsApp.Properties.Settings.PatientsDBConnectionString"].ConnectionString;
            sqlConnection = new SqlConnection(connectionString);
            FillDataGrid();
        }

        private void FillDataGrid()
        {
            //DataGridColumn column = PatientData.AutoSizeMode;
            string query = "SELECT * FROM Patient";
            SqlCommand sqlCommand = new SqlCommand(query, sqlConnection);
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand);
            using (sqlDataAdapter)
            {
                System.Data.DataTable dt = new System.Data.DataTable("Patient");
                sqlDataAdapter.Fill(dt);
                PatientData.ItemsSource = dt.DefaultView;
            }
            //Formatar a Datagrid depois de ser preenchida
            PatientData.Loaded += SetMinWidths;
        }

        //Every column should have minimal width except Discharge that will expand ultil fills window
        public void SetMinWidths(object source, EventArgs e)
        {
            foreach (var column in PatientData.Columns)
            {
                if ((string)column.Header != "Diagnostic")
                    column.MinWidth = column.ActualWidth;
                else
                    column.Width = new DataGridLength(1, DataGridLengthUnitType.Star);
            }
        }


        private void Search_Patient_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string query = "select Id, Name, Age, Gender, Process, Diagnostic, Discharge from Patient";
                query += " where Name like '%' +@name+ '%'";
                query += " and (Age like '%' +@age+ '%')";
                query += " and (Gender like '%' +@gender+ '%')";
                query += " and (Process like '%' +@process+ '%')";
                query += " and (Diagnostic like '%' +@diagnostic+ '%')";
                query += " and (Discharge like '%' +@discharge+ '%')";

                SqlCommand sqlCommand = new SqlCommand(query, sqlConnection);
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand);

                using (sqlDataAdapter)
                {
                    sqlCommand.Parameters.AddWithValue("@name", NameBox.Text);
                    sqlCommand.Parameters.AddWithValue("@age", AgeBox.Text);
                    sqlCommand.Parameters.AddWithValue("@gender", GenderBox.Text);
                    sqlCommand.Parameters.AddWithValue("@process", ProcessBox.Text);
                    sqlCommand.Parameters.AddWithValue("@diagnostic", DiagnosticBox.Text);
                    sqlCommand.Parameters.AddWithValue("@discharge", DischargeBox.Text);

                    System.Data.DataTable dt = new System.Data.DataTable();
                    sqlDataAdapter.Fill(dt);
                    PatientData.ItemsSource = dt.DefaultView;
                }
            }
            catch (Exception ex)
            {
                if (!(ex is MessageBoxException))
                    MessageBox.Show(ex.ToString());
                sqlConnection.Close();
            }
        }

        private void Reset_Patient_Click(object sender, RoutedEventArgs e)
        {
            FillDataGrid();
        }

        private void Save_Patient_Data_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                foreach (DataRowView row in PatientData.ItemsSource)
                {
                    string query = "update Patient set Name=@addName,";
                    query += " Age=@addAge,";
                    query += " Gender=@addGender,";
                    query += " Process=@addProcess,";
                    query += " Diagnostic=@addDiagnostic,";
                    query += " Discharge=@addDischarge";
                    query += " where id = @patientId";
                    SqlCommand sqlCommand = new SqlCommand(query, sqlConnection);
                    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand);

                    using (sqlDataAdapter)
                    {
                        sqlConnection.Open();
                        sqlCommand.Parameters.AddWithValue("@patientId", row.Row.ItemArray[0]);
                        sqlCommand.Parameters.AddWithValue("@addName", row.Row.ItemArray[1]);
                        sqlCommand.Parameters.AddWithValue("@addAge", row.Row.ItemArray[2]);
                        sqlCommand.Parameters.AddWithValue("@addGender", row.Row.ItemArray[3]);
                        sqlCommand.Parameters.AddWithValue("@addProcess", row.Row.ItemArray[4]);
                        sqlCommand.Parameters.AddWithValue("@addDiagnostic", row.Row.ItemArray[5]);
                        sqlCommand.Parameters.AddWithValue("@addDischarge", row.Row.ItemArray[6]);
                        sqlCommand.ExecuteScalar();
                        sqlConnection.Close();
                    }
                }
                MessageBox.Show("Alterações guardadas.");
            }
            catch (Exception ex)
            {
                if (!(ex is MessageBoxException))
                    MessageBox.Show(ex.ToString());
                sqlConnection.Close();
            }
        }
        /* Creating and writing Word .doc */
        private void Export_Data_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = app.Documents.Add();
            try
            {
                foreach (DataRowView row in PatientData.ItemsSource)
                {
                    app.Selection.TypeText("Utente: " + row.Row.ItemArray[1] + "\n");
                    app.Selection.TypeText("Idade: " + row.Row.ItemArray[2] + "\n");
                    app.Selection.TypeText("Genero: " + row.Row.ItemArray[3] + "\n");
                    app.Selection.TypeText("Processo: " + row.Row.ItemArray[4] + "\n");
                    app.Selection.TypeText("\nDiagnostico:\n" + row.Row.ItemArray[5] + "\n");
                    app.Selection.TypeText("\nAlta:\n" + row.Row.ItemArray[6] + "\n");
                    //insert a page break after the last word
                    app.Selection.TypeText("\f");
                }
            }
            catch (Exception ex)
            {
                if (!(ex is MessageBoxException))
                    MessageBox.Show(ex.ToString());
                sqlConnection.Close();
            }

            app.Documents.Save();
            app.Quit();
            MessageBox.Show("Exportado com exito.");
        }
    }
}

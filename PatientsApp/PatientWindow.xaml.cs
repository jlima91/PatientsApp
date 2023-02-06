using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
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
using System.Configuration;
using static System.Net.Mime.MediaTypeNames;

namespace PatientsApp
{
    /// <summary>
    /// Interaction logic for PatientWindow.xaml
    /// </summary>
    public partial class PatientWindow : Window
    {
        readonly SqlConnection sqlConnection;
        readonly ListBox nameList;
        public PatientWindow(ListBox NameList)
        {
            InitializeComponent();
            string connectionString = ConfigurationManager.ConnectionStrings["PatientsApp.Properties.Settings.PatientsDBConnectionString"].ConnectionString;
            sqlConnection = new SqlConnection(connectionString);
            nameList = NameList; //Deve ser feio
            ShowFields(NameList);
        }

        private void ShowFields(ListBox NameList)
        {
            try
            {
                string query = "select * from Patient where id = @patientId";
                SqlCommand sqlCommand = new SqlCommand(query, sqlConnection);
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand);

                using (sqlDataAdapter)
                {
                    sqlCommand.Parameters.AddWithValue("@patientId", NameList.SelectedValue);
                    DataTable dt = new DataTable();
                    sqlDataAdapter.Fill(dt);

                 
                    NameBox.Text = dt.Rows[0]["Name"].ToString();
                    AgeBox.Text = dt.Rows[0]["Age"].ToString();
                    GenderBox.Text = dt.Rows[0]["Gender"].ToString();
                    ProcessBox.Text = dt.Rows[0]["Process"].ToString();
                    DiagnosticBox.Text = dt.Rows[0]["Diagnostic"].ToString();
                    DischargeBox.Text = dt.Rows[0]["Discharge"].ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void Save_Patient_Data_Click(object sender, RoutedEventArgs e)
        {
            int outputValue = 0;
            bool isNumber = int.TryParse(AgeBox.Text, out outputValue);
            try
            {
                if (!isNumber)
                    throw new MessageBoxException("Idade tem de ser um número.");

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
                    sqlCommand.Parameters.AddWithValue("@patientId", nameList.SelectedValue);
                    sqlCommand.Parameters.AddWithValue("@addName", NameBox.Text);
                    sqlCommand.Parameters.AddWithValue("@addAge", AgeBox.Text);
                    sqlCommand.Parameters.AddWithValue("@addGender", GenderBox.Text);
                    sqlCommand.Parameters.AddWithValue("@addProcess", ProcessBox.Text);
                    sqlCommand.Parameters.AddWithValue("@addDiagnostic", DiagnosticBox.Text);
                    sqlCommand.Parameters.AddWithValue("@addDischarge", DischargeBox.Text);
                    sqlCommand.ExecuteScalar();
                    sqlConnection.Close();
                    MessageBox.Show("Alterações guardadas.");
                }
            }
            catch (Exception ex)
            {
                if (!(ex is MessageBoxException))
                    MessageBox.Show(ex.ToString());
                sqlConnection.Close();
            }
        }
        private void Exit_Patient_Data_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}

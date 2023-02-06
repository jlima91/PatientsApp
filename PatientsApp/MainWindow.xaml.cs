using System;
using System.Collections.Generic;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using static System.Net.Mime.MediaTypeNames;
using Microsoft.Office.Interop.Word;

namespace PatientsApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        SqlConnection sqlConnection;
        public MainWindow()
        {
            InitializeComponent();
            string connectionString = ConfigurationManager.ConnectionStrings["PatientsApp.Properties.Settings.PatientsDBConnectionString"].ConnectionString;
            sqlConnection = new SqlConnection(connectionString);

            ShowNames();
        }
        private void ShowNames()
        {
            try
            {
                string query = "select Name, Id from People";
                // The SqlDataAdapter can be imagined like an Interface to make Tables usable by C#-Objects
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(query, sqlConnection);

                using (sqlDataAdapter)
                {
                    System.Data.DataTable nameTable = new System.Data.DataTable();
                    //The Fill method implicitly opens the Connection that the DataAdapter is using if it finds that the connection is not already open. Works for Fill and Update
                    sqlDataAdapter.Fill(nameTable);

                    //Which information of the table in DataTable should be  shown in our nameList
                    NameList.DisplayMemberPath = "Name";
                    //Which value should be delivered when an Item from from our nameList is selected
                    NameList.SelectedValuePath = "Id";
                    //The Reference to the Data the nameList should populate
                    NameList.ItemsSource = nameTable.DefaultView;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void Delete_Patient_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string query = "delete from Patient where id = @patientId";
                SqlCommand sqlCommand = new SqlCommand(query, sqlConnection);
                if (NameList.SelectedValue == null)
                    throw new MessageBoxException("Selecione um utente para remover.");
                sqlConnection.Open();
                sqlCommand.Parameters.AddWithValue("@patientId", NameList.SelectedValue);
                sqlCommand.ExecuteScalar();
                sqlConnection.Close();
                ShowNames();
            }
            catch (Exception ex)
            {
                if (!(ex is MessageBoxException))
                    MessageBox.Show(ex.ToString());
                sqlConnection.Close();
            }
        }

        private void Add_Patient_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string query = "insert into Patient (Name) values (@addName)";
                SqlCommand sqlCommand = new SqlCommand(query, sqlConnection);
                if (string.IsNullOrWhiteSpace(AddNameBox.Text))
                    throw new MessageBoxException("Introduza o nome do utente para adicionar.") ;
                sqlConnection.Open();
                sqlCommand.Parameters.AddWithValue("@addName", AddNameBox.Text);
                sqlCommand.ExecuteScalar();
                sqlConnection.Close();
                ShowNames();
            }
            catch (Exception ex)
            {
                if(!(ex is MessageBoxException))
                    MessageBox.Show(ex.ToString());
                sqlConnection.Close();
            }
        }
        
        private void Search_Patient_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string query = "select Name, Id from Patient";
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

                    NameList.DisplayMemberPath = "Name";
                    NameList.SelectedValuePath = "Id";
                    NameList.ItemsSource = dt.DefaultView;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                sqlConnection.Close();
            }
        }

        private void Reset_Patient_Click(object sender, RoutedEventArgs e)
        {
            ShowNames();
            AddNameBox.Text = "";
        }

        private void NameList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ShowSelectedName();
        }

        private void ShowSelectedName()
        {
            try
            {
                string query = "select Name from Patient where id = @patientId";
                SqlCommand sqlCommand = new SqlCommand(query, sqlConnection);
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand);

                using (sqlDataAdapter)
                {
                    sqlCommand.Parameters.AddWithValue("@patientId", NameList.SelectedValue);
                    System.Data.DataTable patientDataTable = new System.Data.DataTable();
                    sqlDataAdapter.Fill(patientDataTable);

                    AddNameBox.Text = patientDataTable.Rows[0]["Name"].ToString();
                }
            }
            catch (Exception ex) {  } //Retirei porque se nao ocorre excepºao quando se faz reset. o nome é apagado e a funçao selected tenta mostar nome que nao tem id
        }

        private void Edit_Patient_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (NameList.SelectedValue == null)
                    throw new MessageBoxException("Selecione um utente para editar.");
                PatientWindow p = new PatientWindow(NameList);
                p.Show();
            }
            catch (Exception ex) 
            {
                if (!(ex is MessageBoxException))
                    MessageBox.Show(ex.ToString());
            }

        }

        private void Grid_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                GridWindow p = new GridWindow();
                p.Show();
            }
            catch (Exception ex)
            {
                if (!(ex is MessageBoxException))
                    MessageBox.Show(ex.ToString());
            }
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (NameList.SelectedValue == null)
                    throw new MessageBoxException("Selecione um utente para exportar.");

                //Create a Word app
                Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
                // Add a page
                Microsoft.Office.Interop.Word.Document doc = app.Documents.Add();

                string query = "select Name, Age, Gender, Process, Diagnostic, Discharge";
                query += " from Patient where Id = @patientId";
                SqlCommand sqlCommand = new SqlCommand(query, sqlConnection);
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand);

                using (sqlDataAdapter)
                {
                    //Pass to @patientId the id of the selected value
                    sqlCommand.Parameters.AddWithValue("@patientId", NameList.SelectedValue);
                    System.Data.DataTable dt = new System.Data.DataTable();
                    sqlDataAdapter.Fill(dt);
                    app.Selection.TypeText("Utente: " + dt.Rows[0]["Name"].ToString() + "\n");
                    app.Selection.TypeText("Idade: " + dt.Rows[0]["Age"].ToString() + "\n");
                    app.Selection.TypeText("Genero: " + dt.Rows[0]["Gender"].ToString() + "\n");
                    app.Selection.TypeText("Processo: " + dt.Rows[0]["Process"].ToString() + "\n");
                    app.Selection.TypeText("\nDiagnostico:\n" + dt.Rows[0]["Diagnostic"].ToString() + "\n");
                    app.Selection.TypeText("\nAlta:\n" + dt.Rows[0]["Discharge"].ToString() + "\n");
                }
                //Savig before closing
                app.Documents.Save();
                app.Quit();
                MessageBox.Show("Exportado com exito.");
            }
            catch (Exception ex)
            {
                if (!(ex is MessageBoxException))
                    MessageBox.Show(ex.ToString());
                sqlConnection.Close();
            }
        }
    }
}

public class MessageBoxException : Exception
    {
        public MessageBoxException(string message): base(message)
        {
            MessageBox.Show(message);
        }
    }


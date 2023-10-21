using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using System.Xml.Linq;
using System.Data;
using System.Windows.Forms;

namespace TestTask.Models
{
    public enum StateRecord : int
    {
        Active = 1,
        Closed = 0
    }
    public class Employee
    {
        public static string Name { get; set; }
        public static string TableNumber { get; set; }
        public static string JobTitle { get; set; }
        public static SubDivision Subd { get; set; }
        public static string Email { get; set; }
        public static string PhoneNumber { get; set; }
        public static DateTime HireDate { get; set; }
        public static DateTime DismissDate { get; set; }
        public static StateRecord StateRec { get; set; }

        public Employee()
        {

        }

        public Employee(string name, string tableNumber, string jobTitile, string email, string phoneNumber, DateTime hireDate, DateTime dismissDate, StateRecord stateRec)
        {
            Name = name;
            TableNumber = tableNumber;
            JobTitle = jobTitile;
            Email = email;
            PhoneNumber = phoneNumber;
            HireDate = hireDate;
            DismissDate = dismissDate;
            StateRec = stateRec;
        }

        public Employee(int id, SqlConnection sqlConnection)
        {

            SqlCommand cmd = new SqlCommand($"SELECT * FROM Employee WHERE Id = {id}", sqlConnection);
            SqlDataReader reader = cmd.ExecuteReader();
            int idSubD = 0;
            while (reader.Read())
            {
                Name = reader["Name"].ToString();
                TableNumber = reader["Number"].ToString();
                JobTitle = reader["JobTitle"].ToString();
                Email = reader["Email"].ToString();
                PhoneNumber = reader["PhoneNumber"].ToString();
                idSubD = (int)reader["ID_SubDivision"];
                Subd = new SubDivision(idSubD, sqlConnection);
                HireDate = (DateTime)reader["HireDate"];
                Console.WriteLine(HireDate);
                DismissDate = (DateTime)reader["DismissDate"];
                StateRec = (StateRecord)reader["StateRecord"];
                Console.WriteLine(StateRec);
            }
        }

        public void DbInsert(string nameDepartment,SqlConnection sqlConnection)
        {
            SqlDataReader sqlDataReader = null;

            SqlCommand command = new SqlCommand(
                $"INSERT INTO [Employee] (Name, Number, JobTitle, ID_SubDivision, Email, PhoneNumber, HireDate, DismissDate) VALUES (@Name, @Number, @JobTitle, @ID_SubDivision, @Email, @PhoneNumber, @HireDate, @DismissDate)",
                sqlConnection);

            SqlCommand subDivisionCommand = new SqlCommand(
                $"SELECT Id, NameDepartment FROM SubDivision WHERE NameDepartment = @NameDepartment",
                sqlConnection);

            subDivisionCommand.Parameters.AddWithValue("@NameDepartment", nameDepartment);

            try
            {
                sqlDataReader = subDivisionCommand.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(sqlDataReader);
                int numRows = dt.Rows.Count;
                if (numRows > 0)
                {
                    sqlDataReader = subDivisionCommand.ExecuteReader();
                    while (sqlDataReader.Read())
                    {
                        command.Parameters.AddWithValue("@ID_SubDivision", sqlDataReader["Id"]);
                    }
                }
                else
                {
                    subDivisionCommand = new SqlCommand("INSERT INTO [SubDivision] (NameDepartment) VALUES (@NameDepartment)", sqlConnection);

                    subDivisionCommand.Parameters.AddWithValue("@NameDepartment", nameDepartment);
                    subDivisionCommand.ExecuteNonQuery();

                    subDivisionCommand = new SqlCommand("SELECT Id, NameDepartment FROM SubDivision WHERE NameDepartment = @NameDepartment", sqlConnection);
                    subDivisionCommand.Parameters.AddWithValue("@NameDepartment", nameDepartment);

                    sqlDataReader = subDivisionCommand.ExecuteReader();

                    //command.Parameters.AddWithValue("@ID_SubDivision", sqlDataReader["Id"]);

                    while (sqlDataReader.Read())
                    {
                        command.Parameters.AddWithValue("@ID_SubDivision", sqlDataReader["Id"]);
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sqlDataReader?.Close();
            }
            int year = DismissDate.Year == 1 ? 1111 : DismissDate.Year;
            command.Parameters.AddWithValue("@Name", Name);
            command.Parameters.AddWithValue("@Number", TableNumber);
            command.Parameters.AddWithValue("@JobTitle", JobTitle);
            command.Parameters.AddWithValue("@Email", Email);
            command.Parameters.AddWithValue("@PhoneNumber", PhoneNumber);
            command.Parameters.AddWithValue("@HireDate", $"{HireDate.Month}/{HireDate.Day}/{HireDate.Year}");
            command.Parameters.AddWithValue("@DismissDate", $"{DismissDate.Month}/{DismissDate.Day}/{year}");
            command.ExecuteNonQuery();
        }

        public void DbDelete(SqlConnection sqlConnection)
        {
            SqlCommand cmd = new SqlCommand($"DELETE FROM Employee WHERE Number = {TableNumber}", sqlConnection);

            cmd.ExecuteNonQuery();
        }
        public static List<Employee> AllEmployee(SqlConnection sqlConnection)
        {
            SqlCommand cmd = new SqlCommand("SELECT * FROM Employee", sqlConnection);
            SqlDataReader reader = cmd.ExecuteReader();
            List<Employee> listEmp = new List<Employee>();
            while (reader.Read())
            {
                listEmp.Add(new Employee((int)reader["Id"], sqlConnection));
            }

            return listEmp;
        }

        public string ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(Name);
            sb.Append(TableNumber);
            sb.Append(JobTitle);
            sb.Append(Email);
            sb.Append(PhoneNumber);
            sb.Append(HireDate.ToString());
            sb.Append(DismissDate.ToString());
            sb.Append(StateRec.ToString());

            return sb.ToString();
        }
    }
}

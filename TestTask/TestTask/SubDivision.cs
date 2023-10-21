using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using System.Xml.Linq;

namespace TestTask.Models
{
    public class SubDivision
    {
        public int ID { get; set; }
        public string NameDepartment { get; set; }
        public string MainDivisionString { get; set; }
        public SubDivision MainDivision { get; set; }
        public Employee Leader { get; set; }
        public StateRecord StateRecord { get; set; }
        public List<Employee> ListEmployee { get; set; }

        public SubDivision(string nameDepartment, string mainDivision)
        {
            NameDepartment = nameDepartment;
            MainDivisionString = mainDivision;
        }
        public SubDivision(string nameDepartment, SubDivision mainDivision,
                           Employee leader,StateRecord stateRec, List<Employee> listEmployee)
        {
            NameDepartment = nameDepartment;
            MainDivision = mainDivision;
            Leader = leader;
            StateRecord = stateRec;
            ListEmployee = listEmployee;
        }

        public SubDivision(int id, SqlConnection sqlConnection)
        {
            SqlCommand cmd = new SqlCommand($"SELECT * FROM Employee WHERE Id = {id}", sqlConnection);
            SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                ID = (int)reader["Id"];
                NameDepartment = reader["NameDepartment"].ToString();
                MainDivision = GetMainDivision(NameDepartment, sqlConnection);
                Leader = GetLeader(id, sqlConnection);
                StateRecord = (StateRecord)reader["StateRecord"];
                ListEmployee = GetListEmployee(id, sqlConnection);
            }
        }

        private SubDivision GetMainDivision(string nameDepartment, SqlConnection sqlConnection)
        {
            SqlCommand cmd = new SqlCommand("SELECT * FROM SubDivision WHERE NameDepartment = @NameDepartment", sqlConnection);
            cmd.Parameters.AddWithValue("@NameDepartment", nameDepartment);
            SqlDataReader reader = cmd.ExecuteReader();
            SubDivision subDivision = null;
            while (reader.Read())
            {
                subDivision = new SubDivision((int)reader["Id"], sqlConnection);
            }
            return subDivision;
        }

        private List<Employee> GetListEmployee(int id, SqlConnection sqlConnection)
        {
            SqlCommand cmd = new SqlCommand("SELECT * FROM Employee WHERE ID_SubDivision = @ID_SubDivision", sqlConnection);
            cmd.Parameters.AddWithValue("@ID_SubDivision", id);
            SqlDataReader reader = cmd.ExecuteReader();
            List<Employee> listEmp = new List<Employee>();
            while (reader.Read())
            {
                listEmp.Add(new Employee((int)reader["Id"], sqlConnection));
            }

            return listEmp;
        }

        private Employee GetLeader(int id, SqlConnection sqlConnection)
        {
            SqlCommand cmd = new SqlCommand("SELECT * FROM Employee WHERE ID_SubDivision = @ID_SubDivision AND (JobTitle LIKE '%Руководитель%' OR JobTitle LIKE '%Директор%')", sqlConnection);
            cmd.Parameters.AddWithValue("@ID_SubDivision", id);
            SqlDataReader reader = cmd.ExecuteReader();
            int idEmp = 0;
            while (reader.Read())
            {
                idEmp = (int)reader["Id"];
            }
            if(idEmp == 0)
            {
                return null;
            }
            else
            {
                return new Employee(idEmp, sqlConnection);
            }
        }

        public void DbInsert(SqlConnection sqlConnection)
        {
            SqlCommand command = new SqlCommand($"SELECT * FROM SubDivision WHERE NameDepartment = @NameDepartment", sqlConnection);
            command.Parameters.AddWithValue("@NameDepartment", NameDepartment);
            if(command.ExecuteNonQuery() == 0)
            {
                SqlCommand cmd = new SqlCommand($"INSERT INTO [SubDivision] (NameDepartment, MainDivision) VALUES (@NameDepartment, @MainDivision)", sqlConnection);

                cmd.Parameters.AddWithValue("@NameDepartment", NameDepartment);
                cmd.Parameters.AddWithValue("@MainDivision", MainDivisionString);

                cmd.ExecuteNonQuery();
            }
        }

        public void DbDelete(SqlConnection sqlConnection)
        {
            SqlCommand cmd = new SqlCommand($"DELETE FROM Employee WHERE Id = {ID}", sqlConnection);

            cmd.ExecuteNonQuery();
        }

        public static List<SubDivision> AllSubdivision(SqlConnection sqlConnection)
        {
            SqlCommand cmd = new SqlCommand("SELECT * FROM SubDivision", sqlConnection);
            SqlDataReader reader = cmd.ExecuteReader();
            List<SubDivision> listSubD = new List<SubDivision>();
            while (reader.Read())
            {
                listSubD.Add(new SubDivision((int)reader["Id"], sqlConnection));
            }

            return listSubD;
        }
    }
}

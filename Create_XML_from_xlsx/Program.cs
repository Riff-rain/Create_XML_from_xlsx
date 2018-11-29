using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Xml.Linq;

namespace Create_XML_from_xlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            string file_route = "E:\\Source\\Grimoire_Party_Simulator\\Grimoire_Party_Simulator\\Grimoire.xlsx";
            XElement xml = new XElement("Root");

            try
            {
                // OLEDB를 이용해 Excel 데이터를 불러온다.
                String connectionString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = "
                                        + file_route + "; Extended Properties = \"Excel 12.0 Xml;HDR=YES;IMEX = 1\";";

                // Microsoft에 있는 코드 일부 차용.
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    // The insertSQL string contains a SQL statement that
                    // inserts a new row in the source table.
                    OleDbCommand command = new OleDbCommand("Select * From [카드 일람$]", connection);
                    // Open the connection and execute the insert command.
                    try
                    {
                        connection.Open();
                        OleDbDataReader reader = command.ExecuteReader();

                        while (reader.Read())
                        {
                            xml.Add(new XElement("card",
                                    new XElement("preface", reader[0].ToString()),
                                    new XElement("name", reader[1].ToString()),
                                    new XElement("rare", reader[2].ToString()),
                                    new XElement("original_cost", reader[5].ToString()),
                                    new XElement("cost", reader[6].ToString()),
                                    new XElement("attack", reader[7].ToString()),
                                    new XElement("defense", reader[8].ToString()),
                                    new XElement("stat_per_cost", reader[9].ToString()),
                                    new XElement("skill", reader[11].ToString()),
                                    new XElement("skill_arousal", reader[12].ToString()),
                                    new XElement("support", reader[13].ToString()),
                                    new XElement("support_effect", reader[14].ToString()),
                                    new XElement("assist", reader[15].ToString()),
                                    new XElement("assist_effect", reader[16].ToString())
                                ));
                        }
                        reader.Close();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    // The connection is automatically closed when the
                    // code exits the using block.
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            Console.WriteLine(xml);
            xml.Save("Cards.xml");
        }
    }
}

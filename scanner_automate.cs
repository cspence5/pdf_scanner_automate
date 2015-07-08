using System;
using System.IO;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices.ComTypes;
using Oracle.DataAccess.Client;
using Oracle.DataAccess.Types;
using System.Globalization;

namespace rename_file_automation {
	class Program {
		static void Main(string[] args)

		{
			string conString = "User Id=x; password=x;" + "Data Source=x.world; Pooling=false;";




			OracleConnection con = new OracleConnection();
			con.ConnectionString = conString;







			try {
				con.Open();
				Console.WriteLine("Connection to database successful!");
			} catch {
				Console.WriteLine("Connection Failed!");

			}




			string path_string_failed = "";

			string path_string = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));
			int counter = 0;
			path_string_failed = path_string + "\\failed_2\\";

			string[] fileEntries = Directory.GetFiles(path_string_failed, "*.pdf");
			foreach(string fileName in fileEntries) {

				Console.WriteLine(fileName);




				Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
				object miss = System.Reflection.Missing.Value;
				string path_string_z = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @fileName);
				
				object path = path_string_z;
				object readOnly = true;
				Microsoft.Office.Interop.Word.Document docs = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
				string totaltext = "";
				for (int i = 0; i < 100; i++) {

					try {

						totaltext += " \r\n " + docs.Paragraphs[i + 1].Range.Text.ToString();


					} catch {



					}



				}

				int starting_pos = 0;

				int ending_pos = 0;

				try {


					starting_pos = totaltext.IndexOf("Pkt") + 11;
					ending_pos = totaltext.IndexOf("Pkt", starting_pos);



				} catch {

					string newName_path_name_2 = path_string + "\\uhoh\\";
					Console.WriteLine("Bad File Type");




				}


				string pkt_ctrl_num = "";
				string div = "";

				try {


					pkt_ctrl_num = totaltext.Substring(starting_pos, ending_pos - starting_pos).Trim();

					if (pkt_ctrl_num.Length > 13) {

						ending_pos = totaltext.IndexOf("Customer", starting_pos);


						pkt_ctrl_num = totaltext.Substring(starting_pos, ending_pos - starting_pos).Trim();
						pkt_ctrl_num = pkt_ctrl_num.Trim();
						//success = true;

						pkt_ctrl_num = pkt_ctrl_num.Remove(pkt_ctrl_num.Length - 2);

			

					}


				} catch {

					Console.WriteLine("pk:" + pkt_ctrl_num);



					string newName_path_name_2 = path_string + "\\uhoh\\";


					if (pkt_ctrl_num.Length > 12 || pkt_ctrl_num == "") {

						Console.WriteLine("No Pkt Found!");


					} else {
						Console.WriteLine("Incorret File Type...Skipped.");
						System.IO.File.Move(fileName, newName_path_name_2);
						

					}

				}

				Console.WriteLine("Pickticket_found:" + pkt_ctrl_num);



				string query = @
				"select div from pkt_hdr ph 
                            JOIN cd_master cm on ph.cd_master_id = cm.cd_master_id where pkt_ctrl_nbr =" + "'" + pkt_ctrl_num + "'" +
					"and cm.co = 'AMBU'" + @
				"UNION all
                            select div from arc_wh_ghc1.pkt_hdr ap 
                            JOIN cd_master cm on ap.cd_master_id = cm.cd_master_id 
                            where ap.pkt_ctrl_nbr =" + "'" + pkt_ctrl_num + "'" + @
				"
                            and cm.co = 'AMBU' and 
                            not exists ( select null from pkt_hdr ph 
                            JOIN cd_master cm on ph.cd_master_id = cm.cd_master_id 
                            where ph.pkt_ctrl_nbr = " + "'" + pkt_ctrl_num + "'" + ")";

			

				OracleCommand cmd = con.CreateCommand();
				cmd.CommandText = query;

				OracleDataReader reader = cmd.ExecuteReader();

				while (reader.Read()) {
					// Console.WriteLine("in reader");

					div = reader.GetString(0);

					Console.WriteLine("div= " + div);

				}


				string newName = div + "_" + pkt_ctrl_num;
				string newName_path_name = path_string + "\\success2\\" + newName + ".pdf";
				string newName_path_name_3 = path_string + "\\uhoh\\";
				reader.Close();


				docs.Close(false);
				word.Quit(false);
				reader.Close();

				if (docs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(docs);
				if (word != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(word);

				docs = null;
				word = null;
				GC.Collect();


				try {
					// success = false;
					System.IO.File.Move(fileName, newName_path_name);


				} catch {
					// success = false;
					Console.WriteLine("file already exists skipped! ");

					File.Move(fileName, newName_path_name_3 + Path.GetFileName(fileName));
					// System.IO.File.Copy(fileName, newName_path_name_3);

				}



				counter = counter + 1;
			}



			Console.WriteLine("Total Numbers of Files in directory equals: " + counter);
			Console.ReadLine()



		}
	}

}
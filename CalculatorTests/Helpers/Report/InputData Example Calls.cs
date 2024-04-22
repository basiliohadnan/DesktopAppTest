//using System;

//namespace Starline
//{
//    class Program
//    {
//        static void Main(string[] args)
//        {
//            int rows = 0;

//            //AutoReport autoReport = new AutoReport();

//            // PostgreSQL
//            InputData inputPostgreSQL = new InputData(ConnType: "PostgreSQL", ConnServer: "26.205.60.107", ConnPort: 15432, ConnDatabase: "Star_Prod", ConnUser: "ec2-user", ConnPass: "ec2-user");
//            //InputData inputData = new InputData(ConnType: "PostgreSQL", ConnFile: autoReport.GetAppPath() + "/ConexaoSTR/DadosConexao.str");
//            Console.WriteLine("PostgreSQL");
//            rows = inputPostgreSQL.NewQuery("inputPostgreSQL", "select x from arch.teste");
//            Console.WriteLine("@-" + inputPostgreSQL.GetValue("inputPostgreSQL", "X") + "-@");
//            Console.WriteLine("@-" + inputPostgreSQL.RowCount("inputPostgreSQL").ToString() + "-@");
//            inputPostgreSQL.RunDDL("insert into arch.teste values (@v_txt_x)", false, "v_txt_x: z");
//            rows = inputPostgreSQL.NewQuery("inputPostgreSQL2", "select x from arch.teste");
//            Console.WriteLine("@-" + inputPostgreSQL.GetValue("inputPostgreSQL2", "X", inputPostgreSQL.RowCount("inputPostgreSQL2")) + "-@");
//            Console.WriteLine("");

//            // Oracle
//            InputData inputOracle = new InputData(ConnType: "Oracle", ConnServer: "26.167.215.133", ConnPort: 1521, ConnSID: "ORCL", ConnUser: "massa", ConnPass: "massa");
//            Console.WriteLine("Oracle");
//            rows = inputOracle.NewQuery("inputOracle", "select x from teste");
//            Console.WriteLine("@-" + inputOracle.GetValue("inputOracle", "X") + "-@");
//            Console.WriteLine("@-" + inputOracle.RowCount("inputOracle").ToString() + "-@");
//            inputOracle.RunDDL("insert into teste values (:v_txt_x)", false, "v_txt_x: z");
//            rows = inputOracle.NewQuery("inputOracle2", "select x from teste");
//            Console.WriteLine("@-" + inputOracle.GetValue("inputOracle2", "X", inputOracle.RowCount("inputOracle2")) + "-@");
//            Console.WriteLine("");

//            // SQL Server
//            InputData inputSQLServer = new InputData(ConnType: "SQL Server", ConnServer: "26.167.215.133", ConnPort: 1433, ConnDatabase: "massa", ConnUser: "massa", ConnPass: "massa");
//            Console.WriteLine("SQL Server");
//            rows = inputSQLServer.NewQuery("inputSQLServer", "select x from teste");
//            Console.WriteLine("@-" + inputSQLServer.GetValue("inputSQLServer", "X") + "-@");
//            Console.WriteLine("@-" + inputSQLServer.RowCount("inputSQLServer").ToString() + "-@");
//            inputSQLServer.RunDDL("insert into teste values (@v_txt_x)", false, "v_txt_x: z");
//            rows = inputSQLServer.NewQuery("inputSQLServer2", "select x from teste");
//            Console.WriteLine("@-" + inputSQLServer.GetValue("inputSQLServer2", "X", inputSQLServer.RowCount("inputSQLServer2")) + "-@");
//            Console.WriteLine("");

//            // MySQL
//            InputData inputMySQL = new InputData(ConnType: "MySQL", ConnServer: "26.205.60.107", ConnPort: 13306, ConnDatabase: "input_data", ConnUser: "root", ConnPass: "Starline@123");
//            Console.WriteLine("MySQL");
//            rows = inputMySQL.NewQuery("inputMySQL", "select x from teste");
//            Console.WriteLine("@-" + inputMySQL.GetValue("inputMySQL", "X") + "-@");
//            Console.WriteLine("@-" + inputMySQL.RowCount("inputMySQL").ToString() + "-@");
//            inputMySQL.RunDDL("insert into teste values (@v_txt_x)", false, "v_txt_x: z");
//            rows = inputMySQL.NewQuery("inputMySQL2", "select x from teste");
//            Console.WriteLine("@-" + inputMySQL.GetValue("inputMySQL2", "X", inputMySQL.RowCount("inputMySQL2")) + "-@");
//            Console.WriteLine("");

//            // Excel
//            InputData inputExcel = new InputData(ConnType: "Excel", ConnXLS: "C:/Temp/massa.xlsx");
//            Console.WriteLine("Excel");
//            rows = inputExcel.NewQuery("inputExcel", "select x from [teste$]");
//            Console.WriteLine("@-" + inputExcel.GetValue("inputExcel", "X") + "-@");
//            Console.WriteLine("@-" + inputExcel.RowCount("inputExcel").ToString() + "-@");
//            inputExcel.RunDDL("insert into [teste$] values (@v_txt_x)", false, "v_txt_x: z");
//            rows = inputExcel.NewQuery("inputExcel2", "select x from [teste$]");
//            Console.WriteLine("@-" + inputExcel.GetValue("inputExcel2", "X", inputExcel.RowCount("inputExcel2")) + "-@");
//            Console.WriteLine("");

//            // Exemplos de Query
//            /*
//            rows = inputData.NewQuery("cred_oracle", "select auto.get_conf('marisa', 'prm_oracle_hom_ip') as oracle_ip," +
//                                                           " auto.get_conf('marisa', 'prm_oracle_hom_port') as oracle_port");
//            string oracle_ip = inputData.GetValue("cred_oracle", "oracle_ip");
//            string oracle_port = inputData.GetValue("cred_oracle", "oracle_port");
//            rows = inputData.NewQuery("filial", "select auto.get_conf('marisa', @v_txt_prm_name) as filial", "v_txt_prm_name: prm_filial_com_biometria");
//            rows = inputData.NewQuery("filial", "select auto.get_conf('marisa', 'prm_filial_com_biometria') as filial");
//            string filial_nro = inputData.GetValue("filial", "filial");
//            Console.WriteLine(filial_nro);

//            rows = inputData.NewQuery("chart", "select * from auto.confs where cst_id = @v_num_cst_id", "v_num_cst_id: 2");
//            Console.WriteLine(rows.ToString());
//            Console.WriteLine(inputData.QueryData["chart"][1]["conf_value"]);
//            Console.WriteLine(inputData.GetValue("chart", "conf_value"));
//            Console.WriteLine(inputData.GetValue("chart", "conf_value", 2));
//            for (int i = 1; i <= rows; i++)
//            {
//                Console.WriteLine(inputData.QueryData["chart"][i]["conf_value"]);
//            }
//            inputData.rowCount("abc");
//            inputData.getValue("abc", 1, "campox");
//            inputData.runDDL("insert blablabla");
//            */

//            // Exemplos de Base64
//            /*
//            string imageBase64 = "";
//            imageBase64 = oracleHomData.GetImageBase64("select c.usu_bio_imagem as imagem from ccm.ccm_cli_log_bio_fac a, ccm_biometria_facial b, ccm_usuario_certiface c where a.bio_protocolo = b.bio_protocolo and b.id_biometria = c.id_biometria and c.TIPO = 'ENVIADO' and b.cli_cpf = 11504810481");
//            imageBase64 = oracleDevData.GetImageBase64("select foto from starline_base64", "begin" +
//            " insert into starline_base64(foto) select c.usu_bio_imagem from ccm.ccm_cli_log_bio_fac@CCMSTB a," +
//            " ccm_biometria_facial@CCMSTB b, ccm_usuario_certiface@CCMSTB c" +
//            " where a.bio_protocolo = b.bio_protocolo and b.id_biometria = c.id_biometria" +
//            " and c.tipo = 'ENVIADO' and b.cli_cpf = 39256335883;" +
//            " commit;" +
//            " end;");
//            */

//            // Exemplos de DDL
//            /*
//            oracleData.RunDDL("insert into ccm.starline_sql_capture (id) values ('a')");
//            inputData.RunDDL("insert into auto.teste values (@v_txt_x) returning x", true, "v_txt_x: c");
//            */
//        }
//    }
//}

//using System;
//using System.Collections;
//using System.Collections.Generic;
//using System.IO;
//using System.Text;
//using System.Data;
//using System.Data.OleDb;
//using System.Xml;
//using System.Web;


//namespace ConvertToExcel {
//    class ExportToExcel
//    {
//        private static DataTable GetDataTable()
//        {
//            DataTable sampleTable = new DataTable("SampleTable");

//            sampleTable.Columns.Add("Name", typeof(string));
//            sampleTable.Columns.Add("Symbol", typeof(string));
//            sampleTable.Columns.Add("Desc", typeof(int));

//            DataRow dr = sampleTable.Rows.Add();
//            dr[0] = "FirstName";
//            dr[1] = "plus";
//            dr[2] = "description";

//            return sampleTable;
//        }
//    }

//}

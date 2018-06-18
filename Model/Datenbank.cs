using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.ObjectModel;

namespace Model
{
    public class Datenbank
    {
        List<Projekt> lstDoku;
        OleDbConnection con = new OleDbConnection();
        OleDbDataReader reader;

        public List<Projekt> LstDoku
        {
            get
            {
                return lstDoku;
            }

            set
            {
                lstDoku = value;
            }
        }

        private void OpenConnection()
        {
            if (con.State != System.Data.ConnectionState.Open)
            {
                con.Open();
            }
        }

        public Datenbank()
        {
            con = new OleDbConnection(Properties.Settings.Default.DbString);
            LstDoku = new List<Projekt>();
        }

        public List<Projekt> openDatabase()
        {
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandText = "Select * from Files";
            OpenConnection();
            reader = cmd.ExecuteReader();
            int i = 0;
            while(reader.Read())
            {
                Projekt wd = mkDokumentObject(reader, i);
                LstDoku.Add(wd);
            }
            reader.Close();
            con.Close();
            return lstDoku;
        }

        private Projekt mkDokumentObject(OleDbDataReader reader, int i)
        {
            Projekt wd = new Projekt();
            try
            {
                wd.Id = Convert.ToInt32(pruefen(reader[i++]));
                wd.Name = Convert.ToString(pruefen(reader[i++]));
                wd.Objekt = pruefen(reader[i++]);
                wd.Groesse = Convert.ToInt32(pruefen(reader[i++]));
                wd.ErstellDatum = Convert.ToDateTime(pruefen(reader[i++]));
            }
            catch(Exception)
            {
                
            }
            return wd;
        }

        private Object pruefen(Object o)
        {
            if (o == DBNull.Value)
            {
                return null;
            }
            else
            {
                return o;
            }
        }

        private object DokumentenOeffnen(String dokuName)
        {
            foreach(Projekt p in LstDoku)
            {
                if(p.Name.Equals(dokuName))
                {
                    return p;
                }
            }
            return null;
        }

        protected void WordAnzeigen()
        {
            lstDoku.Find(x => x.Dateiart.Contains("word"));

        }


    }
}

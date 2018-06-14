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
        ObservableCollection<Projekt> lstDoku;
        OleDbConnection con = new OleDbConnection();
        OleDbDataReader reader;

        public ObservableCollection<Projekt> LstDoku
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
            con = new OleDbConnection(Model.Properties.Settings.Default.DbString);
            LstDoku = new ObservableCollection<Projekt>();
        }

        public ObservableCollection<Projekt> openDatabase()
        {
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandText = "Select * from Files";
            OpenConnection();
            reader = cmd.ExecuteReader();
            int i = 0;
            while(reader.Read())
            {
                Projekt wd = mkWordDokumentObject(reader, i);
                lstDoku.Add(wd);
            }
            reader.Close();
            con.Close();
            return lstDoku;
        }

        private Projekt mkWordDokumentObject(OleDbDataReader reader, int i)
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
            foreach(Projekt p in lstDoku)
            {
                if(p.Name.Equals(dokuName))
                {
                    return p;
                }
            }
            return null;
        }



    }
}

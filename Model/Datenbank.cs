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
            lstDoku = new List<Projekt>();
        }

        public List<Projekt> openDatabase()
        {
            OpenConnection();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandText = "Select * from Files";
            reader = cmd.ExecuteReader();
            int i = 0;
            while(reader.Read())
            {
                Projekt wd = mkDokumentObject(reader, i);
                lstDoku.Add(wd);
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
                wd.Dateiart = Convert.ToString(pruefen(reader[i++]));
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

        public void delete(Object o)
        {
            try
            {
                Projekt p = (Projekt)o;
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandText = "Delete from Files where Name = " + p.Name;
                OpenConnection();
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception)
            {

                throw;
            }
        }

        public List<Projekt> WordAnzeigen()
        {
            try
            {
                List<Projekt> Liste = new List<Projekt>();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandText = "Select * from Files where dateiart = 'word'";
                OpenConnection();
                reader = cmd.ExecuteReader();
                int i = 0;
                while (reader.Read())
                {
                    Projekt wd = mkDokumentObject(reader, i);
                    Liste.Add(wd);
                }
                reader.Close();
                con.Close();
                return Liste;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public List<Projekt> PowerpointAnzeigen()
        {
            try
            {
                List<Projekt> Liste = new List<Projekt>();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandText = "Select * from Files where dateiart = 'powerpoint'";
                OpenConnection();
                reader = cmd.ExecuteReader();
                int i = 0;
                while (reader.Read())
                {
                    Projekt wd = mkDokumentObject(reader, i);
                    Liste.Add(wd);
                }
                reader.Close();
                con.Close();
                return Liste;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public List<Projekt> ExcelAnzeigen()
        {
            try
            {
                List<Projekt> Liste = new List<Projekt>();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandText = "Select * from Files where dateiart = 'excel'";
                OpenConnection();
                reader = cmd.ExecuteReader();
                int i = 0;
                while (reader.Read())
                {
                    Projekt wd = mkDokumentObject(reader, i);
                    Liste.Add(wd);
                }
                reader.Close();
                con.Close();
                return Liste;
            }
            catch (Exception)
            {

                throw;
            }
        }




    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    public class Projekt
    {
        private int id;
        private String name;
        private object objekt;
        private int groesse;
        private DateTime erstellDatum;
        private String dateiart;

        public Projekt()
        {

        }

        public Projekt(int id, string name, object objekt, int groesse, DateTime erstellDatum, String dateiArt)
        {
            this.id = id;
            this.name = name;
            this.objekt = objekt;
            this.groesse = groesse;
            this.erstellDatum = erstellDatum;
            this.Dateiart = dateiArt;
        }

        public int Id
        {
            get
            {
                return id;
            }

            set
            {
                id = value;
            }
        }

        public string Name
        {
            get
            {
                return name;
            }

            set
            {
                name = value;
            }
        }

        public object Objekt
        {
            get
            {
                return objekt;
            }

            set
            {
                objekt = value;
            }
        }

        public int Groesse
        {
            get
            {
                return groesse;
            }

            set
            {
                groesse = value;
            }
        }

        public DateTime ErstellDatum
        {
            get
            {
                return erstellDatum;
            }

            set
            {
                erstellDatum = value;
            }
        }

        public string Dateiart
        {
            get
            {
                return dateiart;
            }

            set
            {
                dateiart = value;
            }
        }

        public override string ToString()
        {
            return name + " " + erstellDatum + " " + groesse + " " + dateiart;
        }

    }
}

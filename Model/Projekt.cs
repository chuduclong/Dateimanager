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
        private DateTime erstellDatum;
        private String dateiart;

        public Projekt()
        {

        }

        public Projekt(int id, string name, DateTime erstellDatum, String dateiArt, object objekt)
        {
            this.id = id;
            this.name = name;
            this.erstellDatum = erstellDatum;
            this.Dateiart = dateiArt;
            this.objekt = objekt;
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
            return objekt + " " + erstellDatum + " " + dateiart;
        }

    }
}

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
        private DateTime erstellDatum;
        private String dateiart;

        public Projekt()
        {

        }

        public Projekt(int id, string name, DateTime erstellDatum, String dateiArt)
        {
            this.id = id;
            this.name = name;
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
            return name + " " + erstellDatum + " " + dateiart;
        }

    }
}

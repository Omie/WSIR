using System;
using System.Collections.Generic;
using System.Text;

namespace WSIR
{
    class Record
    {
        public String Name { get; set; }
        public String URL { get; set; }

        public Record(String _name, String _url)
        {
            Name = _name;
            URL = _url;
        }

    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ImportSEFData
{
   public  class DatabaseConnection
   {
      
        public string ServerName { get; set; }
        public string DBName { get; set; }
        public string ServerPWD { get; set; }
        public string GlobalConnectionString { get; set; }


    }
}

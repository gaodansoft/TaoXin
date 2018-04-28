using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace Man
{ 
    public class BOM
    {
        public string Name { get; set; }
        public string ID { get; set; }
        public string WName { get; set; }
        public List<BOM> Son = new List<BOM>();


    }
}

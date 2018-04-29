using Newtonsoft.Json;
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
        public string Node { get; set; }
        [JsonIgnore]
        public BOM PBom { get; set; }
        public List<BOM> Son = new List<BOM>();
        public override string ToString()
        {
            if (string.IsNullOrEmpty(Node))
            {
                return string.Format("{0}({1})", Name, ID);
            }
            else
            {
                return string.Format("{0}({1})({2})", Name, ID, Node);
            }
        }
        public void AddSon(BOM Bom)
        {
            Bom.PBom = this;
            Son.Add(Bom);
        }
        public void Remove()
        {
            if (PBom != null)
            {
                PBom.Son.Remove(this);
            }
        }

    }
}

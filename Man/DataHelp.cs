using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Man
{
    public class DataHelp
    {
        public static List<User> UserAll = new List<User>();
        public static List<BOM> RootBom = new List<BOM>();
        public DataHelp()
        {
            for (int i = 0; i < 1000; i++)
            {
                User user = new User() { Id = i.ToString(), Name = "名字" + i.ToString() };
                UserAll.Add(user);
            }
        }
        public List<User> GetUser(string key)
        {
            return UserAll.Where(p => p.Name.Contains(key)).ToList();
        }
        public List<User> GetUserForID(string key)
        {
            return UserAll.Where(p => p.Name.Contains(key)).ToList();
        }
        public List<BOM> GetRootBom()
        {

            //  for (int i = 0; i < 100; i++)
            //{
            //    BOM bom = new BOM(){ID="ID"+i,Name="name"+i};
            //    BOM sbom = new BOM(){ID="sID"+i,Name="sname"+i};
            //    bom.Son.Add(sbom);
            //    RootBom.Add(bom);
            //}
            string text = File.ReadAllText(".\\Data.txt");
            RootBom = JsonConvert.DeserializeObject<List<BOM>>(text);
            return RootBom;


        }
        public void Save()
        {
            string data = JsonConvert.SerializeObject(RootBom);
            File.WriteAllText(".\\Data.txt", data);
        }
    }
}

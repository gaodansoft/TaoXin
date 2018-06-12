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
        public static string DataPath = ".\\Data\\data.txt";
        public static List<User> UserAll = new List<User>();
        public static List<BOM> RootBom = null;
        public static HashSet<string> hs = new HashSet<string>();
        public DataHelp()
        {
            CountData();
            StringBuilder sb = new StringBuilder();
            GetData();
            string[] lines=  File.ReadAllLines(".\\桃一男户口本.csv");
            foreach (var item in lines)
            {
              string[] sl=  item.Split(',');
                if (sl.Length > 1)
                {
                    //  UserAll.Add(new User() { Id = sl[1], Name = sl[0] });
                    if (!hs.Contains(sl[0]))
                    {
                        if(!IsHave(sl[0]))
                           sb.AppendLine(sl[0]);
                    }
                    if (!hs.Contains(sl[1]))
                    {
                        if (!IsHave(sl[1]))
                            sb.AppendLine(sl[1]);
                    }
                }
            }

            var t = sb.ToString();
           
        }

        public bool IsHave(string text)
        {
            foreach (var item in hs)
            {
                if (item.Contains(text))
                    return true;
            }
            return false;
        }
        public HashSet<string> GetData()
        {
           
           
            //HashSet<string> hs = new HashSet<string>();
            string[] lines = File.ReadAllLines(".\\aaaa.csv");
            foreach (var item in lines)
            {
                string[] sl = item.Split(',');
                foreach (var myitem in sl)
                {
                    var d = myitem.Trim();
                    if (!hs.Contains(d))
                    {
                        hs.Add(d);
                    }
                    

                }
            }

            return null;
        }

        public void  CountData()
        {
            StringBuilder sb = new StringBuilder();
            List<CountItem> CountAll = new List<CountItem>();
            try
            {
                //HashSet<string> hs = new HashSet<string>();
                string[] lines = File.ReadAllLines(".\\aaaa.csv");

                for (int i = 0; i < lines.Length; i++)
                {
                    string[] sl = lines[i].Split(',');
                    if (sl.Length >= 14)
                    {
                        if (sl[1].Contains("第一代"))
                        {
                            CountItem cItem = new CountItem();
                            Dictionary<string, string> tempDic = new Dictionary<string, string>();
                            for (int b = 0; b < 100; b++)
                            {
                                var t = lines[i + b + 1].Split(',');
                                if (t.Length < 15) continue;
                               
                                if (t[1].Contains("第一代"))
                                {
                                    CountAll.Add(cItem);
                                    sb.AppendLine(string.Format("{0}\t{2}", cItem.Name, cItem.Count, tempDic.Count));
                                    if (cItem.Count != tempDic.Count)
                                    {

                                    }
                                    
                                    break;
                                }

                                if (string.IsNullOrEmpty(cItem.Name) && t[1].Length >= 2)
                                {

                                    cItem.Name = t[1].First().ToString();
                                }
                                //if (t[18].Length > 0)
                                //{
                                //    cItem.Count += int.Parse(t[18]);
                                //}
                                 GetNOCount(t, tempDic);

                            }
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                var tt = sb.ToString();
            }
            var ttt = sb.ToString();

            // return null;
        }
        public void GetNOCount(string[] lines, Dictionary<string, string> temp)
        {
           
            foreach (var item in lines)
            {
                var t = item.Trim();
                if (item.Length > 17)
                {
                    if (!temp.ContainsKey(t))
                    {
                        temp.Add(t, "");
                    }
                }
            }
           // return temp.Count;
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
            if (RootBom == null)
            {
                if (File.Exists(DataPath))
                {
                    string text = File.ReadAllText(DataPath);
                    RootBom = JsonConvert.DeserializeObject<List<BOM>>(text);
                    return RootBom;
                }
                else
                {
                    RootBom = new List<BOM>();
                }
                
            }
            return RootBom;


        }
        public void Save()
        {

           string dir= Path.GetDirectoryName(DataPath);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            if (File.Exists(DataPath))
            {

                string newPath = Path.Combine(dir, DateTime.Now.ToString("yyyy-MM-dd-hh-mm-ss") + ".txt");
                File.Move(DataPath, newPath);
            }
            string data = JsonConvert.SerializeObject(RootBom);
            File.WriteAllText(DataPath, data);
        }
    }
    class CountItem
    {
        public string Name;
        public int Count;
        public int RCount;
    }
}

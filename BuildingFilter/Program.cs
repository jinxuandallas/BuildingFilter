// See https://aka.ms/new-console-template for more information
using BuildingFilter;
using Everything;
using System.Text;

//var compath = @"txt\com.txt";
var streetpath = @"txt\街道.txt";
var buildingaddresspath = @"txt\楼宇地址.txt";

List<string> streets = ReadStreets(streetpath);
Dictionary<string, string> buildingAddress = ReadBuildingAddress(buildingaddresspath);

//StreamReader sr = new StreamReader(compath);

//string? line = sr.ReadLine();

string txt = ConvertToTxt(GetNewestCompanyFile());
StringReader sr = new StringReader(txt);
string? line = sr.ReadLine();


bool greater;
List<List<string>> content = new List<List<string>>();
//content.Add(new string[] { "企业名称", "企业注册地", "所在楼宇", "企业注册资本(万元)", "统一社会信用代码", "联系电话", "行业门类", "所属管辖街道" }.ToList());
DateTime createDate=DateTime.Now;

while (line != null)
{
    if (line.Contains("注销企业共"))
        break;
    List<string> item = new List<string>();

    if(line.Contains("新注册企业共"))
    {
        string[] t=line.Split('，');
        createDate = DateTime.Parse(t[0]);
    }
    if (line.Contains("："))
    {
        line = line.Trim('。');//去掉句号
        greater = false;
        string[] single = line.Split('，');

        foreach (string s in single)
        {

            string[] kv = s.Split('：');
            if (kv.Length < 2)
                continue;
            if (kv[0].Contains("企业注册资本"))
            {
                float num;
                float.TryParse(kv[1], out num);
                float capital = num;
                if (capital < 500)
                    break;
                else
                    greater = true;
            }

            //新增企业没有成立日期
            //if (kv[0].Contains("成立日期"))
            //    createDate = true;

            item.Add(kv[1].Trim());
        }



        if (greater && streets.Contains(item[item.Count - 1]))//在10个街道中
        {
            //if (!createDate)//如果没有成立日期则填充空白
            //    item.Insert(item.Count - 2, "");
            item.Add(createDate.ToString("d"));
            content.Add(item);
        }

    }
    line = sr.ReadLine();

}
sr.Close();
//Console.WriteLine(line);


Console.WriteLine();
foreach (string s in streets)
    Console.WriteLine(s);

string buildingadd;
//添加楼宇信息
foreach (List<string> item in content)
{
    buildingadd = GetItemAddress(item);
    item.Insert(2, buildingAddress.ContainsKey(buildingadd) ? buildingAddress[buildingadd] : "");
}

///累加不需要添加表头
//content.Insert(0, new string[] { "企业名称", "企业注册地", "所在楼宇", "企业注册资本(万元)", "统一社会信用代码", "联系电话", "成立日期", "行业门类", "所属管辖街道" }.ToList());

OperateExcel oe = new OperateExcel();
//oe.test();
//oe.OperateContent(content);
oe.OperateContentAccumulation(content);


List<string> ReadStreets(string path)
{
    List<string> result = new List<string>();
    StreamReader sr = new StreamReader(path);
    string? line = sr.ReadLine();
    while (line != null)
    {
        result.Add(line.Trim() + "街道");
        line = sr.ReadLine();
    }
    sr.Close();
    return result;

}

Dictionary<string, string> ReadBuildingAddress(string path)
{
    Dictionary<string, string> result = new Dictionary<string, string>();

    StreamReader sr = new StreamReader(path);
    string? line = sr.ReadLine();
    while (line != null)
    {
        string[] item = line.Split('，');
        result.Add(item[1].Trim(), item[0].Trim());
        line = sr.ReadLine();
    }
    sr.Close();

    return result;

}

string GetItemAddress(List<string> item)
{
    int qu = item[1].IndexOf("区");
    int hao = item[1].IndexOf("号", qu);
    return hao < qu ? "" : item[1].Substring(qu + 1, hao - qu).Replace(" ", "");
}

string ConvertToTxt(string path)
{
    Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read);
    Encoding encode = System.Text.Encoding.GetEncoding("gb2312");
    StreamReader sr = new StreamReader(fs, encode);

    string txt = sr.ReadToEnd().Replace("<strong>", string.Empty).Replace(@"</strong>", string.Empty).Replace(@"<br />", "\n").Replace("<html><head></head><body>", string.Empty).Replace("</body></html>", string.Empty);

    //注册企业
    //string aa= sr.ReadToEnd();
    //return aa.IndexOf("注册企业").ToString();
    //return aa;

    return txt;
}

string GetNewestCompanyFile()
{
    EverythingAPI everythingAPI = new EverythingAPI();
    var results = everythingAPI.SearchSortByDate("备注  html !lnk", 1);
    string result=string.Empty;
    foreach (var item in results)
        result = item.ToString();
    return result.Split('\t')[0];
}

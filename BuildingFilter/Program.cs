// See https://aka.ms/new-console-template for more information
using BuildingFilter;

var compath = @"txt\com.txt";
var streetpaht = @"txt\街道.txt";

List<string> streets = ReadStreets(streetpaht);
StreamReader sr = new StreamReader(compath);

string? line = sr.ReadLine();
bool greater, createDate;
List<List<string>> content = new List<List<string>>();
while (line != null)
{
    List<string> item = new List<string>();

    if (line.Contains("："))
    {
        line = line.Trim('。');//去掉句号
        greater = createDate = false;
        string[] single = line.Split('，');

        foreach (string s in single)
        {

            string[] kv = s.Split('：');
            if (kv.Length < 2)
                continue;
            if (kv[0].Contains("企业注册资本"))
            {
                float capital = float.Parse(kv[1]);
                if (capital < 500)
                    break;
                else
                    greater = true;
            }

            if (kv[0].Contains("成立日期"))
                createDate = true;

            item.Add(kv[1].Trim());
        }



        if (greater && streets.Contains(item[item.Count - 1]))//在10个街道中
        {
            if (!createDate)//如果没有成立日期则填充空白
                item.Insert(item.Count - 2, "");
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

OperateExcel oe = new OperateExcel();
//oe.test();
oe.OperateContent(content);


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
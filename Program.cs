using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace DailyReport_CSharp
{
    class Program
    {
        const bool isExcelFormat = true;    // 在实际需求中，输出的文件应该是适合于使用excel打开的样式
        const string DAILY_REPORT_CONFIG = "KT_DaliyReport_Config.txt";
        const int DATE_FORMAT_NUMBER = 4;

        static void Main(string[] args)
        {
			Console.WriteLine("----------------------------------------------------------------");
			Console.WriteLine("欢迎使用 6v 编写的日报统计小程序");
			Console.WriteLine("Version 1.2.1 bulid 20140106");
			Console.WriteLine("");
			Console.WriteLine("最近更新内容：");
			Console.WriteLine("1.2.1　增加了日期样式");
			Console.WriteLine("1.2.0　将结尾的分号'；'替换为句号'。'");
			Console.WriteLine("----------------------------------------------------------------");

			bool isDebug = false;	// 若程序运行时有参数，则认为是启用了测试模式。

            // 1.获取当前日期。这里获取了之后通常没有决定前补0的效果，需要另外处理
            string[] dates = new string[DATE_FORMAT_NUMBER];
            dates[0] = DateTime.Now.ToString("yyyy-MM-dd");
            dates[0] = dates[0].Replace("-", "/");
            dates[1] = DateTime.Now.ToString("yyyy-M-d");
            dates[1] = dates[1].Replace("-", "/");
			dates[2] = DateTime.Now.ToString("yyyy-MM-d");
			dates[2] = dates[2].Replace("-", "/");
			dates[3] = DateTime.Now.ToString("yyyy-M-dd");
			dates[3] = dates[3].Replace("-", "/");

            if (args.Length > 0)
            {
                dates[0] = args[0];
                dates[1] = dates[0].Replace("/0", "/");							// 将2013/05/03这样的日期转为2013/5/3 
				dates[2] = dates[0].Remove(dates[0].LastIndexOf("/0") + 1, 1);	// 将2013/05/03这样的日期转为2013/05/3 
				dates[3] = dates[0].Remove(dates[0].IndexOf("/0") + 1, 1);		// 将2013/05/03这样的日期转为2013/5/03 
				isDebug = true;
            }

            // 2.读取配置文件并预处理
            string configStr = "";
            try
            {
                configStr = File.ReadAllText(DAILY_REPORT_CONFIG);
            }
            catch{
                Console.WriteLine("请在本程序同目录下建立名为＂KT_DailyReport_Config.txt＂的文件");
                return;
            }

            // 将 \r\n 与 \r 这些换行符统一为 \n
            configStr = configStr.Replace("\r\n", "\n");
            configStr = configStr.Replace("\r", "\n");  

            // 3.读取输出文件路径并创建输出流
            string resultFilePath = configStr.Substring(0,configStr.IndexOf("\n"));
            configStr = configStr.Remove(0, resultFilePath.Length + 1);
            FileStream resultFile = new FileStream(resultFilePath, FileMode.Create);
            StreamWriter resultSW = new StreamWriter(resultFile, Encoding.GetEncoding("GB2312"));
            resultSW.WriteLine(dates[0]);
            resultSW.WriteLine(" ");

            while (configStr.IndexOf("\n") != -1)
            {
                // 截取最上一行。
                string currentLine = configStr.Substring(0, configStr.IndexOf("\n"));
                configStr = configStr.Replace(currentLine + "\n", "");

                // 解析当前行内容
                if (currentLine.IndexOf(";") == -1)
                {
                    resultSW.WriteLine(currentLine);
                    continue;
                }
                string nameStr = currentLine.Substring(0, currentLine.IndexOf(";"));
                currentLine = currentLine.Replace(nameStr + ";", "");
                string pathStr = currentLine.Substring(0, currentLine.IndexOf(";"));
                currentLine = currentLine.Replace(pathStr + ";", "");
                string encodeStr = currentLine.Substring(0, currentLine.Length);

                // 确定文件编码
                Encoding encode;
                if (encodeStr.Equals("UCS2LE"))
                    encode = Encoding.Unicode;
                else if (encodeStr.Equals("UTF8"))
                    encode = Encoding.UTF8;
                else if (encodeStr.Equals("ANSI"))
                    encode = Encoding.Default;
                else
                    encode = Encoding.Default;

                // 读取文件内容
                string fileStr = File.ReadAllText(pathStr, encode);
                fileStr = fileStr.Replace("\r\n", "\n");
                fileStr = fileStr.Replace("\r", "\n");

                // 寻找当日日期日期
                string topStr = fileStr.Substring(0, fileStr.IndexOf("\n\n"));
                string dateStr = fileStr.Substring(0, topStr.IndexOf("\n"));
                topStr = topStr.Replace(dateStr + "\n", "");

                // 为一些不合法的格式进行处理
                StringBuilder sb = new StringBuilder(topStr);
                sb.Replace(".\n", "。\n");
                sb.Replace(";\n","。\n");
				sb.Replace("；\n", "。\n");
                sb.Replace(".", "。", sb.Length - 1, 1);
				sb.Replace("；", "。", sb.Length - 1, 1);
                topStr = sb.ToString();


                bool isToday = false;
                for (int i = 0; i < DATE_FORMAT_NUMBER; i++)
                {
                    if (dateStr.IndexOf(dates[i]) != -1)
                    {
                        isToday = true;
                        break;
                    }
                }

				String writeString = "";

                if (isExcelFormat)
                {
                    // 输出
                    if (nameStr.Length == 2)
                    {
                        nameStr = nameStr.Insert(1, "　");
                    }
                    
					writeString += nameStr;

                    if (isToday)
                    {
						writeString = writeString + "\t全勤\t\"" + topStr + "\"\n";
                    }
                    else
                    {
                        if (nameStr.IndexOf("赵随心") != -1)
                        {
							writeString = writeString + "\t全勤\t\"1.参与部门工作。\"\n";
                        }
                        else if (nameStr.IndexOf("李　青") != -1)
                        {
							writeString = writeString + "\t全勤\t\"1.装机器。\"\n";
                        }
                        else
                        {
							writeString = writeString + "\t请假\t\n";
                        }
                    }
                }
                else         // 输出正常的文本格式，非Excel类型
                {
                    // 输出
					writeString = writeString + nameStr + '\n';
                    if (isToday)
                    {
						writeString = writeString + topStr;
                    }
					writeString = writeString + "\n\n";
                }

				resultSW.Write(writeString);
				if (isDebug)
				{
					Console.Write(writeString);
				}
            }

            resultSW.Flush();
            resultSW.Close();
            resultFile.Close();


			if (isDebug)
			{
				Console.Read();
			}
            
        }
    }
}

# dataset to xls
#### 鉴于C#实现dataset数据集导出至excel的坑太多，总结代码如下：

#### 环境：
1. Microsoft Visual C# 2015 
2. Project->Add->Reference->COM->Type Libraries->Microsoft Excel 15.0 Object Library
#### 使用：
* 初始化dataset
* 点击button，dataset数据导入datagridview，然后导出数据至xls文件。（路径是我的文档）
#### 注意：
* 去除格式损坏提示信息
```C# (type)
            //获取你使用的excel 的版本号
            string Version = excel.Version;
            Double FormatNum;
            //使用Excel 97-2003
            if (Convert.ToDouble(Version) < 12)
            {
                FormatNum = -4143;
            }
            //使用excel 2007或者更新de 
            else
            {
                FormatNum = 56;
            } 
```

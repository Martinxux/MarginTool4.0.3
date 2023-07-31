> 问题或建议，请邮件反馈开发人员。

> `xuhui1@cancon.com.cn`
> `niewen@cancon.com.cn`

# 1 项目背景

HGT工具测试后需要耗费很长时间进行数据整理和报告制作且由于 `HG4` 新HGT工具的 log 与 `HG3` 稍微有点不同，为了高效快速制作测试报告，MarginTool 诞生了。经过不断迭代，主要包含以下功能

- 支持 XHCL（XHMI）Margin 的数据分析以及报告生成
- 支持 PCIe 4 & 5 的数据分析以及报告生成
- 支持 PCIe 4 & 5 Preset 测试的所有 Preset 结果提取

# 2 操作指南

## 2.1 安装

直接复制`MarginTool.exe`文件到桌面双击运行即可 *或* 执行`Setup.exe`文件安装

## 2.2 运行

- XHMI 模块
  
   选择测试后的文件夹，注意该文件夹的格式应该如下所示
  
        |-- 父级文件夹
            |-- 8 个子集文件夹（8个 DIE 的结果）
                |-- 5 个子集文件夹（每个 DIE 测试 5 遍）
                    |-- 实际测试文件

***直接选择父级文件夹即可***，会自动对下面所有子文件夹进行分析，再点击生成报告，即在父文件夹目录生成可用的报告

---

- PCIe 模块
  
  选择测试后的文件夹，注意文件夹的格式应该如下所示
  
        |-- 父级文件夹
            |-- 若干个子集文件夹（所有设备的结果）
                |-- 5 个子集文件夹（每个设备测试 5 遍）
                    |-- 实际测试文件

***需要选择第二级子文件夹***，会自动对下面所有文件进行分析，再点击生成报告，即不同设备需要分开分析，对每个设备生成一次报告

---

- PresetScan 模块
  
  选择测试后的文件夹，注意文件夹的格式应该如下所示
  
      |-- 父级文件夹
          |-- 所有 Preset 文件夹（所有 Preset 的结果）
              |-- 实际测试文件

***需要选择父文件夹***，会自动对下面所有文件进行分析，再点击生成报告。

## 2.3 版本控制

2023/07/31

### version: 4.0.3

- 修改界面主题与layout；
- 增加控件图标；
- 增加Suma logo；

---

2023/07/18

### version: 4.0.2

- 修复 Preset 结果判定条件不生效的 BUG；
- 增加了 Preset 中 fom 参数结果；
- 在 WriteToexcel 判断中增加 continue，防止 re 到不同字符引起的崩溃
- 更改打包方式，优化运行速度（用安装包）

---

2023/07/17

### version: 4.0.1

- 修复 PCIe 结果判定只有 PCIe4 生效的 BUG；
- 增加了 PCIe 报告表头、表格式。
- 增加安装包文件 install.exe

---

2023/06/28

### version: 4.0.0

- 合并 XHMI、PCIe Margin4&5、PCIe Preset 的功能，一个 App 三种功能；

- PCIe 模块可以快速提取出五遍的最小值； 

- XHMI 模块可以快速提取 8 个 DIE 的五遍数据的最小值； 

- PresetScan 模块可以提取出 HGT 测试的所有的 Preset。

- 将三个 app 合并并且重新设计 UI 界面，更改名称为 MarginTool。

---

2023/05/23

### version: 2.1.0

- 优化代码结构，不必每个 Die 都输入一次，直接输入父文件夹路径即可，原本需要八次操作，现在只需一次

- 变更报告生成方式，减少内存占用，新增生成报告按钮，选择路径后会将路径下所有子文件夹中的 excel 合并

---

2023/05/10

### version: 1.2.4

- 优化代码结构，将之前分析和制作报告的代码合并

- 增加判断结果功能，生成报告中若有 Fail 值则为红色字体

---

2023/04/26

### version: 1.2.1

- 对 excel 合并代码进行修改，生成可用报告，不必再复制重新做报告

- 修改代码提升运行速度

---

2023/04/07

### version: 1.1.1

- 调整数据、表格格式，增加 CPU 和 DIE 信息、Spec 信息、Worse lane 值

- 优化代码，目录下无需有多余 excel

---

2023/03/25

### version: 1.0.1

- 'Die.xlsx'和'result总表.xlsx'和'AllExcel'切记不能删除或移动！本路径下所有文件必须在同一路径下才能使用

- 将文件进行更名，更符合 XHMI 测试

- 修复了 result 总表中 lane 名显示不全，位置不正确的问题，添加了参数说明

- 只需要运行 XHMI_Anly.exe 即可，运行完毕之后关闭程序，会自动运行 Organize_data.exe，最终结果在 result 总表中查看

---

2023/03/24 

### version: 1.0.0

- 'Result.xlsx'和'result总表.xlsx'和'AllExcel'切记不能删除或移动！

- 将'result总表.xlsx'、'Result.xlsx'、'XHMI_Data_extraction.exe'、'Organize_data.exe'、'AllExcel'所有文件在同一路径下才能运行

- 先打开 XHMI_Data_extraction.exe 进行数据提取，该程序每次只能提取一个DIE，此时会在AllExcel里生成对应的数据，提取八次生成八个数据

- 数据提取完成后，运行 Organize_data.exe ，进行数据汇总，最终数据生成在 result 总表.xlsx 中

## 2.4 引用参考

> [Qt Material](https://github.com/UN-GCPDS/qt-material)

> [Qss](https://blog.csdn.net/y281252548/article/details/109637693)

> [窗口角标](https://www.cnblogs.com/jingsupo/p/13536449.html)

> [QtAwesome](https://github.com/spyder-ide/qtawesome)

> [auto-py-to-exe](https://pypi.org/project/auto-py-to-exe/)

> [HM NIS Edit](https://www.cnblogs.com/yply/p/12001813.html) 

> [nuitka](https://github.com/Nuitka/Nuitka)

## 2.5 图片

![image](https://github.com/xmartin1026/MarginTool4.0.3/blob/main/img.png)

# 3 参与人员

* 徐会
* 聂文

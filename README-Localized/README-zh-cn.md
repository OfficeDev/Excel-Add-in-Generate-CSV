# 适用于 Excel 2016 的 CSV 生成器任务窗格外接程序示例

_适用于：Excel 2016_

此任务窗格外接程序显示了如何使用 Excel 2016 中的 JavaScript API 从列名称的列表创建表。它有两种类型：代码编辑器和 Visual Studio。

![CSV 生成器示例](../Images/ScreenCap1.PNG)

## 尝试一下
### 代码编辑器版本

部署和测试外接程序最简单的方法是将文件复制到网络共享。

1.  在网络共享上创建一个文件夹（如 \\\MyShare\Excel_CSV_Generator），然后复制代码编辑器文件夹中的所有文件。 
2.  编辑清单文件的 <SourceLocation> 元素，使其指向步骤 1 中创建的共享位置。 
3.  将清单 (TeacherCSVGenerator.xml) 复制到网络共享（如 \\\MyShare\MyManifests）。
4.  添加将清单作为 Excel 中受信任的应用目录的共享位置。

    a.启动 Excel 并打开一个空白的电子表格。  
    
    b.选择**文件**选项卡，然后选择**选项**。
    
    c.选择**信任中心**，然后选择**信任中心设置**按钮。
    
    d.选择**受信任的外接程序目录**。
    
    e.在**目录 Url**框中，输入在第 3 步中创建的网络共享的路径，然后选择**添加目录**。
    
   f.  选中“**显示在菜单中**”复选框，然后选择“**确定**”。此时，系统会显示一条消息，提醒你注意你的设置将在 Office 下次启动时应用。 
        
5.  测试并运行外接程序。 

    a.在 Excel 2016 的**插入选项卡**上，选择**我的外接程序**。 
    
    b.在**Office 外接程序**对话框中，选择**共享文件夹**。
    
    c.选择**教师 CSV 班级名单示例**>**插入**。外接程序在任务窗格中打开，并在活动工作表中创建 CSV 班级名单，如此屏幕截图中所示。 
      
   ![大学预算跟踪程序示例](../Images/ScreenCap2.PNG) 

    d.选择一种课堂管理服务。
    
    e.单击生成名单按钮以在活动工作表中插入空的名单。  
    
      ![大学预算跟踪器示例](../Images/ScreenCap3.PNG) 
      
    f.请单击Excel 导出帮助按钮以了解如何将工作表导出为 .csv 文件。  
  
    
### Visual Studio 版本
1.  将项目复制到本地文件夹，并在 Visual Studio 中打开 TeacherCSVGenerator.sln。
2.  按 F5 生成并部署示例外接程序。Excel 启动并且外接程序会在空白工作簿右侧的任务窗格中打开，如以下屏幕截图中所示。 
        
  ![Excel CSV 生成器示例](../Images/ScreenCap1.PNG) 

3.  从下拉列表中选择在线课堂管理服务
4.  使用“**生成名单**”按钮添加学生名单表，然后查看活动工作表中创建的表。

  ![大学预算跟踪程序示例](../Images/ScreenCap3.PNG) 
5.  通过填写表标题下行中的单元格将学生添加到名单中。
6.  使用 Excel 中的导出功能将工作表另存为 .csv 文件。此文件是导入到你所选服务的正确格式。


### 了解详细信息

在您开发外接程序时，Excel JavaScript API 可以提供更多功能。下面只是其中一些可用资源。 

1.  [Excel 外接程序编程概述](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [适用于 Excel 的代码段资源管理器](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Excel 外接程序代码示例](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md) 
4.  [Excel 外接程序 JavaScript API 参考](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [构建你的第一个 Excel 外接程序](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)


# Word转XML格式

# 功能概述  

此方案可将Word(.doc)格式的文档转换为可编辑的XML(.xml)格式的文档。  

# 目录结构  

# 内容列表  
- [功能概述](#功能概述)
- [目录结构](#目录结构)
- [背景说明](#背景说明)
- [转换流程](#转换流程)
- [环境要求](#环境要求)
- [准备条件](#准备条件)
- [操作步骤](#操作步骤)
- [效果说明](#效果说明)
- [技术支持](#技术支持)
- [贡献方式](#贡献方式)
- [开源协议](#开源协议)  

# 背景说明
Word文档是常见的技术文档之一，但Word文档通常是非结构化的，其内部的内容和样式信息不易被机器读取和处理，也不便于文档的管理、重用和自动化处理。  

为了解决这些问题，我们开发了这个`Word to XML Converter`方案，旨在将Word文档转换为XML格式，以便于使用XML编辑器进行文档的进一步编辑、管理和自动化处理。

# 转换流程

>Word(.docx) —— Docbook(.odt) —— XML(.xml) —— DITA(.dita)

# 环境要求    

- Windows 操作系统  
- Python 3.6 或更高版本  
开发者使用的是Python 3.9.4。 https://www.python.org/ftp/python/3.9.4/python-3.9.4-amd64.exe  
- Java开发工具包 OpenJDK  
- Apache OpenOffice  
- lxml 库

# 准备条件
- ## 安装 `Apache OpenOffice`
>**注意**： 此方案须安装win 32-bit (x86) (EXE)版本。  

https://www.openoffice.org/download/index.html  

- ## 安装Java开发工具包 `OpenJDK`  

> **注意**：此方案须安装32-bits版本。

https://github.com/AdoptOpenJDK/openjdk8-binaries/releases/download/jdk8u282-b08/OpenJDK8U-jdk_x86-32_windows_hotspot_8u282b08.msi

- ## 配置 `Apache OpenOffice`
1. 从开始菜单打开 `OpenOffice Writer`，点击 `Tools>Options`，进入配置界面。

2. 在`Options>OpenOffice` 界面中，点击 `Java>Java options`，勾选 `Use a Java runtime environment`, 确认已安装Java运行时的环境。如下图：  


3. 点击 `OK`， 确认退出。  

- ## 安装 `lxml` 库  
打开 `Command Prompt`，输入如下命令：  

> pip install lxml  
    
# 操作步骤 
1. 下载下列文件并放置在同一个路径下：  

    - Python脚本 `word2xml.py`
    - 文件夹 `XSLs`

2. 将Word文档转换为Docbook格式。  

    用`OpenOffice Writer` 打开一个Word文档，并另存为Docbook格式。  

3. 将Dockbook文档转换为XML格式。  
    
    用文本编辑器 `Notepad++` 打开第(1)步保存的Docbook文档，删除 `<article>`标签以外的信息，将`<article>` 标签及以内的信息保存在`word2xml.py`所在路径，命名为 `input.xml`。    

4. 使用脚本，基于`input.xml`生成 DITA 文档。  

    (1) 打开 `Command Prompt`，用`cd` + `<filepath>` 进入脚本所在路径。  

    (2) 在脚本所在路径，输入以下命令运行脚本，将在同一路径下生成名为 `output.dita` 的文件：

    > python word2xml.py 

    示例如下：  


    
    生成的 `output.dita`文件可在XML编辑器（`DITA CSM` 或`Oxygen XML Editor`）中打开和编辑。  


# 效果说明  
本方案可以帮助用户将现有的Word文档转换为更适合于XML编辑器中编辑和管理的格式，实现以下优势：

- **内容组织与管理更规范**：转换为XML后，文档内容将以一种更结构化和层次化的方式呈现，使得内容的组织和管理更加灵活和规范。
- **内容重用和共享更便捷**：XML格式的文档更容易被拆分可重用的模块，显著减少重复工作、提高文档的一致性，并促进团队之间的内容共享和协作。  
- **自动化处理成为可能**：通过将文档转换为XML，可以使用各种XML处理工具和技术来自动提取、转换、验证和发布文档内容，提高工作效率。  
- **多平台访问不再受限**： 转换为XML格式后，文档可以在各种XML编辑器、内容管理系统（CMS）、翻译工具之间进行无缝交互和共享。  

## 技术支持

-   邮箱：hui.wang@sigmatechnology.com

## 贡献方式

非常欢迎你的加入！你可以通过提一个issue 或 pull request的方式来共同开发。

## 开源协议


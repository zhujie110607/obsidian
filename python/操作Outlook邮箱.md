
# 获取Outlook邮箱指定文件夹中邮件的信息，比如发件人，发送时间，主题，正文，附件等

* 在Python中，你可以使用`pywin32`库来访问并操作Windows上的Outlook邮箱。以下是一个基本的例子，演示如何连接到Outlook，访问指定文件夹，并获取邮件的基础信息：

```python
import win32com.client  
  
# 启动Outlook应用  
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")  
  
# 通过文件夹名获取文件夹，例如收件箱可以使用"6"  
inbox = outlook.GetDefaultFolder(6)  
  
# 在指定文件夹中获取所有邮件  
messages = inbox.Items  
  
for message in messages:  
    try:  
        print("发件人: ", message.SenderName)  
        print("发件人邮箱: ", message.SenderEmailAddress)  
        print("发送时间: ", message.SentOn)  
        print("收件人: ", message.To)  
        print("抄送: ", message.CC)  
        print("主题: ", message.Subject)  
        print("正文: ", message.Body)  
        # 对于附件，需要进行遍历  
        attachments = message.Attachments  
        if attachments.Count > 0:  
            print("附件:")  
            for attachment in attachments:  
                print(attachment.FileName)  
        print("--------------------------")  
    except Exception as e:  
        print("Error: ", e)
```

* 请记住：  
  
需要先安装pywin32库，可以使用pip install pywin32命令安装。  
这个脚本应该在具有Outlook和相应配置文件的Windows系统上运行。  
如果Outlook邮箱中有大量邮件，这段代码可能会需要相当的时间来执行。  
由于安全原因，某些Outlook配置可能不允许脚本访问邮件内容。  
这段代码假定你有权从Outlook中读取邮件。如果你的邮箱配置或公司策略不允许，这段代码可能无法正常工作。  
请根据需要修改文件夹、邮件索引或其他条件来定位你需要的邮件。

# 如果邮件正文是HTML格式，您通常需要解析HTML以获取所需的文如果Outlook邮件的正文是HTML格式的，你可以依然使用`pywin32`模块去获取邮件的HTML正文，然后利用HTML解析库，如`BeautifulSoup`，来解析HTML并提取所需信息。下面的步骤和代码示例将指导你完成这个过程：

首先，确保你已经安装了`beautifulsoup4`，你可以通过运行以下命令来安装：

```python
pip install beautifulsoup4
```
* 然后，你可以通过`pywin32`来获取正文为HTML的邮件，并使用`BeautifulSoup`来解析它：

```python
import win32com.client  
from bs4 import BeautifulSoup  
  
# 连接到Outlook  
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")  
  
# 选择"Inbox"文件夹  
inbox = outlook.GetDefaultFolder(6)  # 6代表“Inbox”  
  
# 过滤或选择特定文件夹下的邮件  
messages = inbox.Items  
  
for message in messages:  
    # 获取邮件属性  
    sender = message.SenderName  
    sent_time = message.SentOn  
    subject = message.Subject  
    # 获取邮件的HTML正文  
    html_body = message.HTMLBody  
  
    # 解析HTML正文  
    soup = BeautifulSoup(html_body, features="html.parser")  
    # 提取纯文本  
    plain_text_body = soup.get_text()  
  
    # 打印信息  
    print("Sender:", sender)  
    print("Sent On:", sent_time)  
    print("Subject:", subject)  
    print("Body:", plain_text_body)  
  
    # 处理附件  
    attachments = message.Attachments  
    for attachment in attachments:  
        print("Attachment:", attachment.FileName)  
        # ...保存或处理附件  
  
    # 请注意，为了测试，你可能只想处理一封邮件  
    break
```

* 这段代码首先使用`pywin32`获取Outlook应用程序和"Inbox"文件夹中的邮件条目。它遍历邮件，提取发件人、发送时间、主题和HTML正文。然后，使用`BeautifulSoup`进行HTML解析，从而提取非HTML文本。这对于输出邮件的可读信息是有帮助的。最后，它还展示了如何遍历和打印附件的名称。

# 要使用Python生成一封含有文本正文、HTML表格和Excel文件附件的Outlook邮件，可以继续使用`pywin32`库。以下是一个示例脚本，展示如何创建一封这样的邮件：

```python
import win32com.client as win32  
from pathlib import Path  
  
# 创建一个Outlook实例  
outlook = win32.Dispatch('outlook.application')  
  
# 创建一封邮件  
mail = outlook.CreateItem(0)  
  
# 邮件主题  
mail.Subject = 'Your subject here'  
  
# 邮件正文（文本）  
mail.Body = 'This is the email body in text format.'  
  
# 邮件正文HTML，你可以按照HTML规则设计你的表格  
mail.HTMLBody = '''  
This is the email body in HTML format.  
  
<table border=1>  
    <tr>        
	    <th>Header 1</th>        
	    <th>Header 2</th>    
	</tr>    
	<tr>        
		<td>Cell 1</td>        
		<td>Cell 2</td>    
	</tr>
</table>  
'''  
  
# 添加附件（确保路径正确）  
attachment_path = Path("your_excel_file.xlsx")  
mail.Attachments.Add(str(attachment_path))  
  
# 显示邮件前预览  
mail.Display(True)
```
* 请确保Excel文件的路径是正确的。你也可以通过`mail.To`，`mail.CC`和`mail.BCC`属性来设置收件人、抄送和密送。如果你想直接发送邮件，可以使用`mail.Send()`，但是请注意，在发送邮件前总是建议先预览邮件，确保所有内容都是正确的。

在运行此脚本之前，请确保你已经安装了pywin32库，如果没有安装，可以通过`pip`来安装：

```python
pip install pywin32
```
* 还需注意，前提是你的计算机上安装了Outlook，并且你有权限使用它发送邮件。
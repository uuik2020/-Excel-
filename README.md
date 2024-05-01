# 二维码生成工具

所有代码由微软 Copilot 生成

这是一个使用Python的Tkinter库创建的二维码生成工具。它的主要功能包括生成二维码和将数据保存到Excel文件中。以下是代码的主要部分的解释：

1. **导入所需的库**：这个部分导入了所有需要的Python库，包括`tkinter`（用于创建图形用户界面），`qrcode`（用于生成二维码），`PIL`（用于处理图片），`pandas`（用于处理数据和保存到Excel文件），`os`（用于处理文件和目录路径）和`datetime`（用于获取当前的日期和时间）。

2. **初始化全局变量**：这个部分初始化了一些全局变量，包括`label`（用于显示二维码图片的标签），`num_label`（用于显示数字的标签）和`num`（用于生成二维码的数字）。

3. **定义生成二维码的函数**：`generate_qr`函数首先删除旧的二维码（如果存在），然后创建一个新的二维码对象，添加数据到二维码对象，生成二维码，保存二维码为图片，最后显示二维码图片和数字。

4. **定义保存数据到Excel的函数**：`save_to_excel`函数首先获取输入的数据和数字，然后指定数据存储路径，如果数据存储路径不存在，则创建。接着指定数据文件路径，如果数据文件存在，则读取数据文件，添加新的数据到数据文件，否则创建新的数据文件。最后保存数据到数据文件，数字加一。

5. **创建主窗口**：这个部分创建了一个主窗口，设置了窗口的标题，指定了数据存储路径和数据文件路径，如果数据文件存在，则读取数据文件中的最后一个数字，否则设置初始数字。然后创建了一个输入框，两个按钮（一个用于生成二维码，一个用于保存数据到Excel），最后运行主循环。


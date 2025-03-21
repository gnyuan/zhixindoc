工具函数
========

本章节列出了所有知心插件提供的工具函数，每个函数均可直接在 Excel 中调用。下文对各工具函数的语法、参数说明、返回值以及使用示例进行了详细说明，帮助用户快速上手使用。

1 生成二维码函数 (zt_Qrcode)
----------------------------
该函数用于根据指定文本生成二维码图片，并支持图片尺寸的多种调整模式。  
**Excel 调用语法：**

.. code-block:: none

    =zt_Qrcode(words, [mode], [height], [width], [picture], [colorized])

**参数说明：**

- **words** (必需)：二维码内容（例如 URL 或文本）。
- **mode** (可选，默认 4)：调整图片大小模式  
  - 1：保持宽高比适应单元格  
  - 2：拉伸/压缩以填满单元格  
  - 3：保持原始尺寸  
  - 4：自定义尺寸（需同时指定 height 和 width）
- **height** (可选，默认 120)：自定义模式下图片高度（像素）。
- **width** (可选，默认 120)：自定义模式下图片宽度（像素）。
- **picture** (可选)：图片绝对路径，用于与二维码合成。
- **colorized** (可选，默认 TRUE)：指定二维码是否彩色。

**返回值：**  
成功时返回描述二维码生成状态的文本；失败时返回错误信息。

**使用示例：**

.. code-block:: none

    =zt_Qrcode("https://www.bilibili.com/video/BV16i421e7Ft")
    =zt_Qrcode("https://example.com", 4, 150, 150)
    =zt_Qrcode("https://example.com", 4, 120, 120, "C:\images\background.jpg", TRUE)


2 在单元格中插入图片函数 (zt_Image)
---------------------------------------
该函数用于在 Excel 单元格中插入指定 URL 的图片。  
**Excel 调用语法：**

.. code-block:: none

    =zt_Image(url, [mode], [height], [width])

**参数说明：**

- **url** (必需)：图片的网址，必须包含协议（例如 http://）。
- **mode** (可选，默认 3)：调整图片大小模式  
  - 1：保持宽高比适应单元格  
  - 2：拉伸/压缩以填满单元格  
  - 3：保持原始尺寸  
  - 4：自定义尺寸（需指定 height 与 width）
- **height** (可选)：自定义模式下图片高度（像素）。
- **width** (可选)：自定义模式下图片宽度（像素）。

**返回值：**  
成功时返回“输出图片”，否则返回错误信息。

**使用示例：**

.. code-block:: none

    =zt_Image("https://www.google.com/images/srpr/logo3w.png")
    =zt_Image("https://www.google.com/images/srpr/logo3w.png", 4, 200, 200)


3 下载视频/音频函数 (zt_Video)
--------------------------------
该函数用于根据提供的 URL 下载视频或音频文件。  
**Excel 调用语法：**

.. code-block:: none

    =zt_Video(url, [media_type], [proxy])

**参数说明：**

- **url** (必需)：视频或音频的 URL 地址。
- **media_type** (可选，默认 "video")：下载类型，"video" 或 "audio"。
- **proxy** (可选)：代理地址（例如 socks5://127.0.0.1:8888）。

**返回值：**  
函数为异步调用，依次返回处理状态、下载进度和结果信息。

**使用示例：**

.. code-block:: none

    =zt_Video("https://www.bilibili.com/video/BV16i421e7Ft/")
    =zt_Video("https://www.bilibili.com/video/BV16i421e7Ft/", "audio")


4 发送邮件函数 (zt_Mail)
---------------------------
该函数用于批量发送邮件，并支持添加附件。  
**Excel 调用语法：**

.. code-block:: none

    =zt_Mail(subject, body, recipients, [filename])

**参数说明：**

- **subject** (必需)：邮件标题。
- **body** (必需)：邮件正文。
- **recipients** (必需)：收件人邮箱地址，可为多个单元格。
- **filename** (可选)：附件的绝对路径。

**返回值：**  
返回邮件发送状态信息。

**使用示例：**

.. code-block:: none

    =zt_Mail("邮件标题", "邮件正文", B7:C7)
    =zt_Mail("邮件标题", "邮件正文", B7:C7, "D:\附件.txt")


5 调用大语言模型函数 (zt_LLM)
------------------------------
该函数调用配置好的大语言模型，根据提示词生成回答。  
**Excel 调用语法：**

.. code-block:: none

    =zt_LLM(prompt, [input], [system_content], [is_show_reasoning], [temperature])

**参数说明：**

- **prompt** (必需)：提示词，可包含占位符 {}。
- **input** (可选)：辅助输入内容，用于替换 prompt 中的 {}。
- **system_content** (可选，默认 "You are a helpful assistant")：系统角色描述。
- **is_show_reasoning** (可选，默认 FALSE)：是否显示推理过程。
- **temperature** (可选，默认 0.7)：生成回答的创意性（0-2）。

**返回值：**  
异步返回大语言模型生成的回答内容。

**使用示例：**

.. code-block:: none

    =zt_LLM("用五岁小孩口吻告诉我量子力学")
    =zt_LLM("把{}翻译成英文", "量子力学")


6 利用示例调用大语言模型函数 (zt_LLMEx)
-------------------------------------------
该函数通过少量示例，使大语言模型更好理解用户需求，适用于多列输入。  
**Excel 调用语法：**

.. code-block:: none

    =zt_LLMEx(examples, inputs, [prompt], [system_content], [is_show_reasoning], [temperature])

**参数说明：**

- **examples** (必需)：两列数据，第一列为输入，第二列为对应输出示例。
- **inputs** (必需)：待生成回答的输入数据（单列）。
- **prompt** (可选)：额外提示词。
- **system_content** (可选)：系统角色描述。
- **is_show_reasoning** (可选)：是否显示推理过程。
- **temperature** (可选)：生成回答的创意性（0-2）。

**返回值：**  
异步返回生成的回答，多个输出以分隔符区分。

**使用示例：**

.. code-block:: none

    =zt_LLMEx(A1:B3, C1:C10, "它是省份和省会的对应关系")


7 生成同类测试数据函数 (zt_LLMGen)
-----------------------------------
该函数根据输入示例数据，生成指定数量的同类型测试数据。  
**Excel 调用语法：**

.. code-block:: none

    =zt_LLMGen(input_data, [num], [system_content], [is_show_reasoning], [temperature])

**参数说明：**

- **input_data** (必需)：一列示例数据，用于确定数据类型。
- **num** (可选，默认 10)：需要生成的数据数量。
- **system_content** (可选)：系统角色描述。
- **is_show_reasoning** (可选)：是否显示推理过程。
- **temperature** (可选)：生成数据的创意性（0-2）。

**返回值：**  
异步返回生成的测试数据列表。

**使用示例：**

.. code-block:: none

    =zt_LLMGen(A1:A3, 2)


8 爬虫函数 (zt_Crawler)
--------------------------
该函数用于抓取指定页面中目标元素的数据。  
**Excel 调用语法：**

.. code-block:: none

    =zt_Crawler(xpath, [sub_xpaths], [is_cumulative], [is_reload], [browser_url])

**参数说明：**

- **xpath** (必需)：目标列表中某元素的 XPath（示例中为第五个元素）。
- **sub_xpaths** (可选)：具体提取数据的 XPath 表达式（多项）。
- **is_cumulative** (可选，默认 FALSE)：是否累计不同数据。
- **is_reload** (可选，默认 FALSE)：是否每次循环刷新页面。
- **browser_url** (可选，默认 "http://localhost:9222")：连接 Chrome 调试接口的 URL。

**返回值：**  
异步返回爬取的数据。

**使用示例：**

.. code-block:: none

    =zt_Crawler("/html/body/div/div/div[2]/div[2]/div[1]/div[2]/div[5]/div")


9 爬虫（自动翻页）函数 (zt_CrawlerAll)
----------------------------------------
该函数支持自动点击“下一页”，对列表数据进行持续爬取。  
**Excel 调用语法：**

.. code-block:: none

    =zt_CrawlerAll(xpath, [next_button], [stop_control_cell], [browser_url])

**参数说明：**

- **xpath** (必需)：目标列表中某元素的 XPath（须包含 [5]）。
- **next_button** (可选)：下一页按钮的 XPath。
- **stop_control_cell** (可选，默认 "A1")：置值停止翻页的控制单元格。
- **browser_url** (可选，默认 "http://localhost:9222")：Chrome 调试接口 URL。

**返回值：**  
异步返回累计爬取的列表数据。

**使用示例：**

.. code-block:: none

    =zt_CrawlerAll("/html/body/div/div/div[2]/div[2]/div[1]/div[2]/div[5]/div", "/html/body/div[3]/div[1]/div/div[1]/div[2]/span[3]/a")


10 生成公钥函数 (zt_GenPubKey)
--------------------------------
该函数使用用户记忆的口令生成公钥，并将生成结果保存至指定目录。  
**Excel 调用语法：**

.. code-block:: none

    =zt_GenPubKey(user_privkey)

**参数说明：**

- **user_privkey** (必需)：用户私钥，用于生成公钥。

**返回值：**  
返回生成状态信息，若已存在则提示用户先删除旧公钥。

**使用示例：**

.. code-block:: none

    =zt_GenPubKey("myPrivateKey")


11 加密函数 (zt_Encrypt)
-------------------------
该函数使用已生成的公钥对明文进行加密。  
**Excel 调用语法：**

.. code-block:: none

    =zt_Encrypt(plaintext)

**参数说明：**

- **plaintext** (必需)：需要加密的明文字符串。

**返回值：**  
返回加密后的密文；若未生成公钥则提示错误。

**使用示例：**

.. code-block:: none

    =zt_Encrypt("需要加密的内容")


12 解密函数 (zt_Decrypt)
-------------------------
该函数使用提供的私钥对密文进行解密。  
**Excel 调用语法：**

.. code-block:: none

    =zt_Decrypt(private_key_str, encrypted_data)

**参数说明：**

- **private_key_str** (必需)：私钥字符串。
- **encrypted_data** (必需)：待解密的密文。

**返回值：**  
返回解密后的明文。

**使用示例：**

.. code-block:: none

    =zt_Decrypt("myPrivateKey", "加密后的内容")


13 文字转语音函数 (zt_tts)
----------------------------
该函数利用 Windows 内置接口将文字转换为语音播放。  
**Excel 调用语法：**

.. code-block:: none

    =zt_tts(text)

**参数说明：**

- **text** (必需)：待转换为语音的文本字符串。

**返回值：**  
异步返回处理状态。

**使用示例：**

.. code-block:: none

    =zt_tts("你好，hello world")


14 发送通知函数 (zt_Notice)
-----------------------------
该函数用于发送桌面通知，可设置通知延迟。  
**Excel 调用语法：**

.. code-block:: none

    =zt_Notice(title, content, [delay_seconds])

**参数说明：**

- **title** (可选)：通知标题。
- **content** (可选)：通知内容。
- **delay_seconds** (可选，默认 3)：延迟发送通知的秒数。

**返回值：**  
返回设置通知成功的提示信息。

**使用示例：**

.. code-block:: none

    =zt_Notice("通知的标题", "通知的内容", 10)

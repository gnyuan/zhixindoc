转换函数
============

本章节介绍知心插件中提供的文本转换工具函数，这些函数主要用于处理文本格式转换以及日期格式转换。下文详细说明每个函数的功能、调用语法、参数及使用示例，帮助用户在 Excel 中高效调用。

1 将文本转换为大写 (zv_to_uppercase)
--------------------------------------
该函数将输入文本中的所有字母转换为大写形式，保留非字母字符不变。

**Excel 调用语法：**

.. code-block:: none

    =zv_to_uppercase(text)

**参数说明：**

- **text** (文本，字符串，默认空串)：待转换为大写的文本。

**返回值：**  
转换后的文本，所有字母均为大写。

**示例：**

.. code-block:: none

    =zv_to_uppercase("hello 123")    // 返回 "HELLO 123"
    =zv_to_uppercase("AbC测试")       // 返回 "ABC测试"


2 将文本转换为小写 (zv_to_lowercase)
--------------------------------------
该函数将输入文本中的所有字母转换为小写形式，保留非字母字符不变。

**Excel 调用语法：**

.. code-block:: none

    =zv_to_lowercase(text)

**参数说明：**

- **text** (文本，字符串，默认空串)：待转换为小写的文本。

**返回值：**  
转换后的文本，所有字母均为小写。

**示例：**

.. code-block:: none

    =zv_to_lowercase("HELLO 123")    // 返回 "hello 123"
    =zv_to_lowercase("AbC测试")       // 返回 "abc测试"


3 将文本首字母大写 (zv_tcapitalize_first_letter)
--------------------------------------------------
该函数仅将文本的首个字母转换为大写，其余字符保持不变。

**Excel 调用语法：**

.. code-block:: none

    =zv_tcapitalize_first_letter(text)

**参数说明：**

- **text** (文本，字符串，默认空串)：待转换的文本。

**返回值：**  
返回转换后的文本，只有首字母大写。

**示例：**

.. code-block:: none

    =zv_tcapitalize_first_letter("hello world")         // 返回 "Hello world"
    =zv_tcapitalize_first_letter("python programming")    // 返回 "Python programming"


4 将每个单词首字母大写 (zv_capitalize_each_word)
----------------------------------------------------
该函数将以空格分隔的每个单词的首字母转换为大写，其余字符保持不变，适用于标题格式化。

**Excel 调用语法：**

.. code-block:: none

    =zv_capitalize_each_word(text)

**参数说明：**

- **text** (文本，字符串，默认空串)：待转换的文本。

**返回值：**  
返回转换后的文本，每个单词的首字母均为大写。

**示例：**

.. code-block:: none

    =zv_capitalize_each_word("hello world")           // 返回 "Hello World"
    =zv_capitalize_each_word("python   programming")    // 返回 "Python   Programming"


5 去除首尾空格 (zv_strip_spaces)
---------------------------------
该函数用于去除文本两端的空格，但保留中间空格不变。

**Excel 调用语法：**

.. code-block:: none

    =zv_strip_spaces(text)

**参数说明：**

- **text** (文本，字符串，默认空串)：待处理的文本。

**返回值：**  
返回去除首尾空格后的文本。

**示例：**

.. code-block:: none

    =zv_strip_spaces("  hello world  ")         // 返回 "hello world"
    =zv_strip_spaces("  python programming  ")    // 返回 "python programming"


6 将文本转换为 Excel 日期 (zv_to_date)
-----------------------------------------
该函数将表示日期的文本转换为 Excel 日期格式，支持多种日期表示方式。

**Excel 调用语法：**

.. code-block:: none

    =zv_to_date(text)

**参数说明：**

- **text** (文本，字符串，默认空串)：待转换为日期的文本，支持的格式包括：
  - 20200101
  - 2020-01-01
  - 2020/01/01

**返回值：**  
返回转换后的 Excel 日期格式（如 2020/1/1）。

**示例：**

.. code-block:: none

    =zv_to_date("20200101")     // 返回 2020/1/1
    =zv_to_date(20200101)       // 返回 2020/1/1
    =zv_to_date("2020-01-01")     // 返回 2020/1/1

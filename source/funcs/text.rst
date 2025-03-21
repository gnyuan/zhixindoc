文本处理函数
================

本章节介绍知心插件中提供的文本处理函数，这些函数可用于验证、提取、匹配、替换、拆分和组合文本数据。

1 ze_IsDate 函数
---------------------
验证输入值是否为有效的日期格式。

**Excel 调用语法：**

.. code-block:: none

    =ze_IsDate(value)

**参数说明：**

- **value**: 需要验证的日期值。

**返回值：**

返回 TRUE（有效日期）或 FALSE（无效日期）。

**示例：**

.. code-block:: none

    =ze_IsDate("2023-10-05")   // 返回 TRUE
    =ze_IsDate("20230230")     // 返回 FALSE

2 ze_IsEmail 函数
---------------------
检查输入的值是否为有效的电子邮件地址。

**Excel 调用语法：**

.. code-block:: none

    =ze_IsEmail(value)

**参数说明：**

- **value**: 需要验证的电子邮件地址。

**示例：**

.. code-block:: none

    =ze_IsEmail("noreply@google.com")   // 返回 TRUE

3 ze_RegexExtract 函数
---------------------------
根据正则表达式提取所有匹配子串，并用逗号分隔。

**Excel 调用语法：**

.. code-block:: none

    =ze_RegexExtract(text, regex, [delimiter])

**参数说明：**

- **text**: 输入文本。
- **regex**: 用于匹配的正则表达式。
- **delimiter**: 分隔多个匹配项的字符串，默认为逗号（,）。

**示例：**

.. code-block:: none

    =ze_RegexExtract("电子123表456格", "\d+")  // 返回 "123,456"

4 ze_RegexMatch 函数
------------------------
判断输入文本是否与正则表达式匹配。

**Excel 调用语法：**

.. code-block:: none

    =ze_RegexMatch(text, regex)

**示例：**

.. code-block:: none

    =ze_RegexMatch("电子表格", "S.r")  // 返回 FALSE

5 ze_RegexReplace 函数
---------------------------
使用正则表达式将文本中的部分内容替换为其他文本。

**Excel 调用语法：**

.. code-block:: none

    =ze_RegexReplace(text, regex, replacement)

**示例：**

.. code-block:: none

    =ze_RegexReplace("电子表格", "S.*d", "床")

6 ze_Split 函数
------------------
按指定分隔符拆分文本，并将结果存放在单元格中。

**Excel 调用语法：**

.. code-block:: none

    =ze_Split(text, delimiter, [split_each_char], [remove_empty_text])

**示例：**

.. code-block:: none

    =ze_Split("1,2,3", ",")
    =ze_Split("Alas, poor Yorick", " ")

7 ze_Join 函数
-----------------
将多个字符串或数组中的文本合并，并使用指定分隔符分隔。

**Excel 调用语法：**

.. code-block:: none

    =ze_Join(delimiter, ignore_empty, text1, text2, ...)

**示例：**

.. code-block:: none

    =ze_Join(" ", TRUE, "hello", "world")  // 返回 "hello world"

8 ze_IsURL 函数
------------------
检查某个值是否为有效的网址。

**Excel 调用语法：**

.. code-block:: none

    =ze_IsURL(value)

**示例：**

.. code-block:: none

    =ze_IsURL("https://www.baidu.com")  // 返回 TRUE
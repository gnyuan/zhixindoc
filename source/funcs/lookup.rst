查找函数
============

本章节介绍知心插件中提供的查找工具函数，这些函数旨在为 Excel 用户提供智能、精准的查找和匹配功能，超越传统的 VLOOKUP/HLOOKUP 方法。

1 智能 VLOOKUP 函数 (zl_NlpLookup)
-------------------------------------
该函数类似于 Excel 中的 VLOOKUP，但利用自然语言处理技术进行分词和查找，从而获得更精准的语义匹配结果。

**Excel 调用语法：**

.. code-block:: none

    =zl_NlpLookup(lookup_value, table_array, [col_index_num], [include_score], [all_matches], [threshold], [tokenizer_type])

**参数说明：**

- **lookup_value**  
  待查找的值（数字或字符串），类似于 VLOOKUP 中的查找值。

- **table_array**  
  数据矩阵，类似于 VLOOKUP 中的查找范围，其中第一列作为关键查找列。

- **col_index_num**  
  指定返回结果所在的列号（非必填）。

- **include_score**  
  是否返回匹配分数，布尔值（默认 False）。

- **all_matches**  
  是否只返回全匹配结果，布尔值（默认 False）。

- **threshold**  
  匹配灵敏度阈值（默认 0.3）；阈值越高，匹配要求越严格。

- **tokenizer_type**  
  中文分词引擎，可选 "jieba" 或 "char"；默认情况下，根据传入文本长度自动选择分词方式。

**返回值：**  
返回一个 PDFrame（数据框），包含查找到的结果，及可选的匹配分数。

**使用示例：**

.. code-block:: none

    =zl_NlpLookup(I4, D:E, 1)


2 XLOOKUP 函数 (zl_XLookup)
-----------------------------
该函数实现了类似 Excel XLOOKUP 的功能，用于在单行或单列范围中查找指定值，并返回相应的结果。它既可以取代 VLOOKUP，也可以取代 HLOOKUP，适用于更灵活的查找场景。

**Excel 调用语法：**

.. code-block:: none

    =zl_XLookup(search_key, lookup_range, result_range, [missing_value], [match_mode], [search_mode])

**参数说明：**

- **search_key**  
  要查找的值。

- **lookup_range**  
  要搜索的范围（必须为单行或单列）。

- **result_range**  
  返回结果所在的范围，其行数或列数应与 lookup_range 相同，取决于查找方式。

- **missing_value**  
  如果未找到匹配值，则返回该值（默认返回 #N/A）。

- **match_mode**  
  查找匹配值的方式（默认 0）。

- **search_mode**  
  搜索 lookup_range 的方式（默认 1）。

**返回值：**  
返回查找到的值；如果未匹配，则返回 missing_value。

**使用示例：**

.. code-block:: none

    =zl_XLookup("Apple", A2:A, E2:E)    // 用于取代 VLOOKUP("Apple", A2:E, 5, FALSE)
    =zl_XLookup("Price", A1:E1, A6:E6)    // 用于取代 HLOOKUP("Price", A1:E6, 6, FALSE)
    =zl_XLookup("Apple", E2:E7, A2:A7)    // 等效于 VLOOKUP("Apple", {E2:E7, A2:A7}, 2, FALSE)

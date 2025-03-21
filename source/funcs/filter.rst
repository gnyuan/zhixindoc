过滤器函数
===============

本章节介绍知心插件中提供的过滤器工具函数，这些函数主要用于对 Excel 数据范围进行排序、筛选和去重。下文详细说明每个函数的功能、调用语法、参数及使用示例，帮助用户快速掌握数据过滤操作。

1 按指定列排序 (zf_Sort)
-------------------------
该函数依据一列或多列中的值对数组或范围中的行进行排序。

**Excel 调用语法：**

.. code-block:: none

    =zf_Sort(excel_range, sort_column, ascending, [sort_column2], [ascending2])

**参数说明：**

- **excel_range**：要排序的数据范围。
- **sort_column**：排序依据所在列的列号（可在范围内或范围之外指定）。
- **ascending**：排序方式，TRUE 表示升序排序，FALSE 表示降序排序。
- **sort_column2** (可选)：用于次级排序的列号，优先级较低。
- **ascending2** (可选)：次级排序的排序方式，TRUE 表示升序，FALSE 表示降序。

**返回值：**  
返回按指定排序规则排列的二维数组。

**使用示例：**

.. code-block:: none

    =zf_Sort(A2, 1, TRUE)           // 按首列升序排序
    =zf_Sort(C2, 3, FALSE, 2, TRUE)   // 优先按第三列降序、第二列升序排序


2 返回排序后前 n 个项目 (zf_SortN)
----------------------------------
该函数对数据范围进行排序后，返回排序后的前 n 个项目。排序规则支持多列排序，且可选择显示与第 n 项相同的其他行。

**Excel 调用语法：**

.. code-block:: none

    =zf_SortN(excel_range, n, mode, [sort_column1], [ascending1], [sort_column2], [ascending2])

**参数说明：**

- **excel_range**：要排序并提取前 n 个项目的数据范围。
- **n**：返回的项目数量，必须大于 0。
- **mode**：显示相同值的模式：
  - 0：最多显示排序范围中的前 n 行。
  - 1：显示前 n 行及与第 n 行相同的其他行。
  - 2：去除重复行后显示前 n 行。
  - 3：显示前 n 个不重复行，但包含这些行的所有重复项。
- **sort_column1**：用于排序的第一列的列号，指定范围必须为单列数据，行数与 excel_range 相同。
- **ascending1**：第一排序列的排序方式，TRUE 表示升序，FALSE 表示降序。
- **sort_column2** (可选)：用于次级排序的列号。
- **ascending2** (可选)：次级排序的排序方式，TRUE 表示升序，FALSE 表示降序。

**返回值：**  
返回排序后数据集中前 n 个项目的二维数组。

**使用示例：**

.. code-block:: none

    =zf_SortN(A1, 2)                // 返回前 2 个最小数值
    =zf_SortN(B2, 3, 1)             // 返回包含与第三项相同值的所有条目


3 返回唯一性过滤结果 (zf_Unique)
-----------------------------------
该函数用于对数据范围进行唯一性过滤，剔除重复行或仅返回出现一次的记录。

**Excel 调用语法：**

.. code-block:: none

    =zf_Unique(excel_range, [by_column], [exactly_once])

**参数说明：**

- **excel_range**：需要过滤的数据范围。
- **by_column** (可选)：布尔值，指示是否按列进行过滤。默认值为 FALSE（按行过滤）。
- **exactly_once** (可选)：布尔值，若为 TRUE 则仅返回在范围中只出现一次的记录；默认值为 FALSE。

**返回值：**  
返回过滤后具有唯一性的二维数组。

**使用示例：**

.. code-block:: none

    =zf_Unique([[1,2],[1,2],[3,4]])           // 返回 [[1,2],[3,4]]
    =zf_Unique(A1, exactly_once=TRUE)           // 仅返回在 A1 范围中仅出现一次的记录

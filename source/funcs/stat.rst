统计函数
==================

本章节介绍知心插件中提供的统计工具函数，包括基于条件筛选的最小值和最大值计算函数。

1 MINIFS 函数 (zs_MinIfs)
----------------------------
该函数返回单元格范围中的最小值，并按一组条件进行过滤，类似于 Excel 中的 MINIFS 函数。

**Excel 调用语法：**

.. code-block:: none

    =zs_MinIfs(excel_range, criteria_range1, criterion1, [criteria_range2, criterion2], ...)

**参数说明：**

- **excel_range**  
  要计算最小值的单元格范围。

- **criteria_range1**  
  用于评估 criterion1 的单元格范围。

- **criterion1**  
  应用于 criteria_range1 的模式或测试。

- **criteria_range2, criterion2, ...（可选）**  
  额外的条件范围及其对应的条件，可以添加多个条件。

**返回值：**  
返回满足所有条件的单元格范围内的最小值。

**使用示例：**

.. code-block:: none

    =zs_MinIfs(A1:A3, B1:B3, 1, C1:C3, "A")
    =zs_MinIfs(D4:E5, F4:G5, ">5", F6:G7, "<10")


2 MAXIFS 函数 (zs_MaxIfs)
---------------------------
该函数返回单元格范围中的最大值，并按一组条件进行过滤，类似于 Excel 中的 MAXIFS 函数。

**Excel 调用语法：**

.. code-block:: none

    =zs_MaxIfs(excel_range, criteria_range1, criterion1, [criteria_range2, criterion2], ...)

**参数说明：**

- **excel_range**  
  要计算最大值的单元格范围。

- **criteria_range1**  
  用于评估 criterion1 的单元格范围。

- **criterion1**  
  应用于 criteria_range1 的模式或测试。

- **criteria_range2, criterion2, ...（可选）**  
  额外的条件范围及其对应的条件，可以添加多个条件。

**返回值：**  
返回满足所有条件的单元格范围内的最大值。如果未满足任何条件，则返回 0。  
如果 `excel_range` 和所有 `criteria_range` 的大小不匹配，则返回 `#VALUE!` 错误。

**使用示例：**

.. code-block:: none

    =zs_MaxIfs(A1:A3, B1:B3, 1, C1:C3, "A")
    =zs_MaxIfs(D4:E5, F4:G5, ">5", F6:G7, "<10")

序列函数
============

本章节介绍知心插件提供的时间序列处理函数，这些函数用于填充缺失值、分解时间序列数据，以及进行时间序列预测。

1 三次样条插值填充 (zs_FillWithCS)
-------------------------------------
该函数使用三次样条插值法填充数列中的缺失值，适用于时间序列数据的平滑填充。

**Excel 调用语法：**

.. code-block:: none

    =zs_FillWithCS(arr)

**参数说明：**

- **arr**  
  目标数列，包含需要填充缺失值的数列。

**返回值：**  
返回填充缺失值后的数列。

**使用示例：**

.. code-block:: none

    =zs_FillWithCS(E2:E8)


2 KNN 插值填充 (zs_FillWithKNN)
------------------------------------
该函数使用 KNN（K 近邻）算法填充缺失值，参考完整数据表格 arr1，根据邻近样本估算目标数列 arr2 的缺失值。

**Excel 调用语法：**

.. code-block:: none

    =zs_FillWithKNN(arr1, arr2, [n_neighbors])

**参数说明：**

- **arr1**  
  参考数列，包含完整数据，用于提供填充参考。

- **arr2**  
  目标数列，包含需要填充缺失值的数列。

- **n_neighbors**  
  KNN 算法中使用的邻居数量，默认为 3。

**返回值：**  
返回填充缺失值后的目标数列。

**使用示例：**

.. code-block:: none

    =zs_FillWithKNN(D2:D8, E2:E8)
    =zs_FillWithKNN(D2:D8, E2:E8, 4)


3 时间序列分解 (zs_SeasonalDecompose)
---------------------------------------
该函数对时间序列数据进行分解，拆分为趋势成分、季节性成分和残差成分，以分析数据的周期性规律。

**Excel 调用语法：**

.. code-block:: none

    =zs_SeasonalDecompose(data, ds, y, [period])

**参数说明：**

- **data**  
  带表头的表格数据。

- **ds**  
  日期轴字段名，字符串类型。

- **y**  
  数据轴字段名，字符串类型。

- **period**  
  季节性周期长度，默认 12。年周期为 365，月周期为 12。

**返回值：**  
返回时间序列的趋势、季节性和残差分解结果。

**使用示例：**

.. code-block:: none

    =zs_SeasonalDecompose(B2:C59, B2, C2)


4 时间序列预测 (zs_Prophet)
-----------------------------
该函数使用 Meta（Facebook）开发的 Prophet 模型进行时间序列预测，适用于具有季节性趋势的数据。

**Excel 调用语法：**

.. code-block:: none

    =zs_Prophet(data, ds, y, [periods])

**参数说明：**

- **data**  
  带表头的表格数据。

- **ds**  
  日期轴字段名，字符串类型。

- **y**  
  数据轴字段名，字符串类型。

- **periods**  
  预测周期，默认 365（年周期）。

**返回值：**  
返回预测的时间序列数据。

**使用示例：**

.. code-block:: none

    =zs_Prophet(B2:C59, B2, C2)

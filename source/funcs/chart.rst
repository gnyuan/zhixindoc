图表函数
============

1 zc_BubbleChart (气泡图)
-------------------------------------
根据表格数据绘制多维气泡图

**Excel 调用语法：**

.. code-block:: none

    =zc_BubbleChart(data, x, y, [bubble_size], [category], [label], [title])

**参数说明：**

- **data**  
  必须选择包含表头的二维表格区域（如"A1:F200"）

- **x**  
  X轴字段名，必须是表头存在的列名（如"销售额"）

- **y**  
  Y轴字段名，必须是表头存在的列名（如"利润")

- **bubble_size** *(可选)*  
  指定气泡大小对应的数值列（如"市场规模"）

- **category** *(可选)*  
  分类字段名，不同类别用颜色区分（如"区域"）

- **label** *(可选)*  
  显示在气泡上的标签字段（如"产品线"）

- **title** *(可选)*  
  图表标题文本（如"销售分析图"）

**示例：**

.. code-block:: none

    =zc_BubbleChart(B2:G150, "X轴", "Y轴", "规模", , "名称")

2 zc_Sankey (桑基图)
---------------------------------
绘制资源流动关系图

**Excel 调用语法：**

.. code-block:: none

    =zc_Sankey(nodes, connections, [title], [browser])

**参数说明：**

- **nodes**  
  节点定义区域（需2列）：
  
  - 第1列：节点名称（如"北美"）  
  - 第2列：节点类型（如"收入"）

- **connections**  
  连接关系区域（至少3列）：
  
  - 源节点名称  
  - 目标节点名称  
  - 连接值  
  - 同比（可选）

- **title** *(可选)*  
  图表标题（如"资金流向图"）

- **browser** *(可选，默认FALSE)*  
  设为TRUE时在浏览器打开交互图表

**示例：**

.. code-block:: none

    =zc_Sankey(A1:B10, D1:G50, , TRUE)

3 zc_Scatter (散点图)
---------------------------------
绘制二维散点图

**Excel 调用语法：**

.. code-block:: none

    =zc_Scatter(data_frame, [query], [browser], kwargs)

**参数说明：**

- **data_frame**  
  包含坐标数据的表格区域（如"A1:E100"）

- **query** *(可选)*  
  数据过滤条件（如"销售额>1000"）

- **browser** *(默认FALSE)*  
  TRUE输出交互式网页图表

- **kwargs**  
  两列参数设置区域，常用参数：
  
  ============  ======================
  参数名        说明                  
  ============  ======================
  x             X轴字段名            
  y             Y轴字段名            
  color         颜色映射字段          
  symbol        形状分类字段          
  ============  ======================

**示例：**

.. code-block:: none

    =zc_Scatter(A1:D200, , TRUE, B1:C3)
    /* kwargs区域内容：
       x    经度
       y    纬度 
       color 城市等级
    */

4 zc_Scatter3d (三维散点图)
--------------------------------------
绘制三维空间散点图

**Excel 调用语法：**

.. code-block:: none

    =zc_Scatter3d(data_frame, [query], [browser], kwargs)

**特殊参数要求：**

- 必须通过kwargs指定z轴字段：

.. code-block:: none

    z    海拔高度

**示例：**

.. code-block:: none

    =zc_Scatter3d(A1:F500, , TRUE, B1:C4)
    /* kwargs：
       x    温度
       y    湿度
       z    气压
    */

5 zc_Line (折线图)
------------------------------
绘制趋势折线图

**Excel 调用语法：**

.. code-block:: none

    =zc_Line(data_frame, [query], [browser], kwargs)

**特殊参数：**

- 通过kwargs设置日期字段：

.. code-block:: none

    x         日期
    y         指标值
    line_dash 线型（solid/dash/dot）

**示例：**

.. code-block:: none

    =zc_Line(A1:C365, "年份=2023", , B1:C3)
    /* kwargs：
       x    日期
       y    销售额
    */

6 c_Bar (条形图)
----------------------------
绘制分类对比条形图

**Excel 调用语法：**

.. code-block:: none

    =zc_Bar(data_frame, [query], [browser], kwargs)

**特殊参数：**

- 需指定分类字段和数值字段：

.. code-block:: none

    x        类别字段
    y        数值字段
    barmode  group/stack（分组/堆叠）

**示例：**

.. code-block:: none

    =zc_Bar(A1:D50, , , B1:C2)
    /* kwargs：
       x    产品类型
       y    销量
    */

7 zc_Pie (饼图)
--------------------------
绘制比例分布饼图

**Excel 调用语法：**

.. code-block:: none

    =zc_Pie(data_frame, [query], [browser], kwargs)

**参数要求：**

- 必须包含分类字段和数值字段：

.. code-block:: none

    names    分类字段
    values   数值字段

**示例：**

.. code-block:: none

    =zc_Pie(A1:B10, , , B1:C2)
    /* kwargs：
       names   省份
       values  人口占比
    */

8 zc_WordCloud (词云图)
----------------------------------
生成文本数据词云

**Excel 调用语法：**

.. code-block:: none

    =zc_WordCloud(data_frame, word_column, weight_column)

**参数说明：**

- **data_frame**  
  包含词语和权重的表格（如"A1:B100"）

- **word_column**  
  词语字段名（默认"词语"）

- **weight_column**  
  权重字段名（默认"权重"）

**示例：**

.. code-block:: none

    =zc_WordCloud(D1:E200, "关键词", "出现频次")

9 zc_Histogram (直方图)
--------------------------------------
绘制数值分布直方图

**Excel 调用语法：**

.. code-block:: none

    =zc_Histogram(data_frame, [query], [browser], kwargs)

**特殊参数：**

- 需指定数值字段：

.. code-block:: none

    x        数值字段
    nbins    分箱数量（默认自动）

**示例：**

.. code-block:: none

    =zc_Histogram(A1:C1000, , , B1:C2)
    /* kwargs：
       x     考试成绩
       nbins 20
    */
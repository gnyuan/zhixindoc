行情函数
============

本章节介绍知心插件中的行情相关函数，这些函数提供实时股票数据查询和图表插入功能。

1 实时行情查询 (zq_Realtime)
------------------------------
该函数用于获取指定股票的实时行情数据。

**Excel 调用语法：**

.. code-block:: none

    =zq_Realtime(symbol, [field])

**参数说明：**

- **symbol**  
  标的新浪代码，例如 `sh000001` 表示上证综指。

- **field** *(可选，默认为 `lastPrice` 最新价)*  
  指定查询的指标，可选值包括：
  
  - `open` - 开盘价  
  - `preClose` - 昨收盘  
  - `high` - 最高价  
  - `low` - 最低价  
  - `volume` - 成交手数  
  - `turnOver` - 成交额  
  - `bidPrice1` / `askPrice1` - 买1/卖1价格  
  - `bidVolume1` / `askVolume1` - 买1/卖1数量  
  - 以及更多新浪行情字段。

**返回值：**  
返回指定股票的实时行情数据。

**使用示例：**

.. code-block:: none

    =zq_Realtime("sh000001")
    =zq_Realtime("sh000001", "open")

2 日 K 线图 (zq_DailyK)
-------------------------
该函数用于在 Excel 中插入指定股票的日 K 线图。

**Excel 调用语法：**

.. code-block:: none

    =zq_DailyK(symbol, [mode], [height], [width])

**参数说明：**

- **symbol**  
  新浪股票代码，如 `sh000001`。

- **mode** *(可选，默认为 3)*  
  图片调整模式：
  
  1. 保持宽高比适合单元格
  2. 拉伸或压缩适合单元格
  3. 保持原始大小
  4. 自定义大小（需提供 `height` 和 `width`）

- **height** / **width** *(可选)*  
  以像素为单位的高度和宽度，仅当 `mode=4` 时有效。

**示例：**

.. code-block:: none

    =zq_DailyK("sh000001")
    =zq_DailyK("sh000001", 4, 300, 500)

3 周 K 线图 (zq_WeeklyK)
--------------------------
该函数用于插入周 K 线图。

**语法：**

.. code-block:: none

    =zq_WeeklyK(symbol, [mode], [height], [width])

**示例：**

.. code-block:: none

    =zq_WeeklyK("sh000001")

4 月 K 线图 (zq_MonthlyK)
---------------------------
该函数用于插入月 K 线图。

**语法：**

.. code-block:: none

    =zq_MonthlyK(symbol, [mode], [height], [width])

**示例：**

.. code-block:: none

    =zq_MonthlyK("sh000001")

5 分钟线图 (zq_MinuteK)
-------------------------
该函数用于插入日内分钟线图。

**语法：**

.. code-block:: none

    =zq_MinuteK(symbol, [mode], [height], [width])

**示例：**

.. code-block:: none

    =zq_MinuteK("sh000001")
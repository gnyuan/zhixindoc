知心插件功能函数
================

知心宏里面包含了一系列操作，确实需要文档来说明说明。但知心函数按道理是不需要额外的文档来描述的，下面的示意图可以看到，文档是自带在函数之中的，调用时实时看到它简洁清晰的描述。但既然这是知心插件官网，还是有必要罗列部分函数的描述：

.. image:: images/formula.gif
   :alt: 知心函数


日期函数
--------

- **这天月份有几天** ``(zd_days_in_month)``: 返回指定日期所在月份的天数。
- **这天年份有几天** ``(zd_days_in_year)``: 返回指定日期所在年份的天数。
- **这天是本周第几天** ``(zd_day_of_week)``: 返回指定日期是所在周的第几天。
- **这天是本月第几周** ``(zd_week_of_month)``: 返回指定日期是所在月份的第几周。
- **这天是本年第几周** ``(zd_week_of_year)``: 返回指定日期是所在年份的第几周。
- **这天是本年第几季** ``(zd_quarter_of_year)``: 返回指定日期是所在年份的第几季度。

日期区间函数
-------------

- **这天加减数字** ``(zr_days_in_month)``: 对指定日期加减指定的天数。
- **这天距离今天有几天** ``(zr_days_until_date)``: 计算从今天到指定日期的天数。
- **两日期间多少天** ``(zr_days_between_dates)``: 计算两个指定日期之间的天数。
- **两日期间多少工作日** ``(zr_working_days_between_dates)``: 计算两个指定日期之间的工作日天数。
- **两日期间多少非工作日** ``(zr_non_working_days_between_dates)``: 计算两个指定日期之间的非工作日天数。

信息提取函数
-------------

- **身份证取年龄** ``(zi_id_card_to_age)``: 根据身份证号码提取年龄。
- **身份证取生日** ``(zi_id_card_to_birthday)``: 根据身份证号码提取生日。
- **身份证取性别** ``(zi_id_card_to_gender)``: 根据身份证号码提取性别。
- **邮箱取用户名** ``(zi_email_to_username)``: 根据邮箱地址提取用户名。
- **邮箱取域名** ``(zi_email_to_domain)``: 根据邮箱地址提取域名。
- **网址取域名** ``(zi_url_to_domain)``: 根据网址提取域名。

数据提取函数
-------------

- **提取数值** ``(ze_extract_numbers)``: 从文本中提取数值。
- **提取电话** ``(ze_extract_phone_numbers)``: 从文本中提取电话号码。
- **提取身份证** ``(ze_extract_id_cards)``: 从文本中提取身份证号码。
- **提取邮箱** ``(ze_extract_emails)``: 从文本中提取邮箱地址。
- **提取超链** ``(ze_extract_hyperlinks)``: 从文本中提取超链接。
- **提取日期** ``(ze_extract_dates)``: 从文本中提取日期。
- **提取时间** ``(ze_extract_times)``: 从文本中提取时间。
- **提取邮编** ``(ze_extract_postal_codes)``: 从文本中提取邮编。
- **提取货币金额** ``(ze_extract_currency_amounts)``: 从文本中提取货币金额。

转换函数
---------

- **全部大写** ``(zc_to_uppercase)``: 将文本转换为全部大写。
- **全部小写** ``(zc_to_lowercase)``: 将文本转换为全部小写。
- **句首字母大写** ``(zc_tcapitalize_first_letter)``: 将文本的句首字母转换为大写。
- **词首字母大写** ``(zc_capitalize_each_word)``: 将文本中每个词的首字母转换为大写。
- **删除左右边界空格** ``(zc_strip_spaces)``: 删除文本左右两侧的空格。

序列函数
---------

- **录入星期** ``(zs_week_input)``: 根据序号返回星期。
- **录入月份** ``(zs_month_input)``: 根据序号返回月份。
- **录入甲乙丙丁** ``(zs_heavenly_stems_input)``: 根据序号返回天干。
- **录入子某寅丑** ``(zs_earthly_branches_input)``: 根据序号返回地支。
- **录入天干地支** ``(zs_heavenly_stems_earthly_branches_cycle)``: 根据序号返回天干地支序列。
- **录入八卦** ``(zs_bagua_input)``: 根据序号返回八卦。
- **录入星座** ``(zs_zodiac_input)``: 根据序号返回星座。
- **录入节气** ``(zs_solar_term_input)``: 根据序号返回节气。
- **录入生肖** ``(zs_chinese_zodiac_input)``: 根据序号返回生肖。
- **录入五行** ``(zs_five_elements_input)``: 根据序号返回五行。

以上是知心插件提供的部分功能函数的介绍，旨在帮助用户提高在Excel中的工作效率和数据处理能力。
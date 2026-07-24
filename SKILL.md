---
name: test-report-generator
description: 从测试用例xls/xlsx文件提取用例数据并按docx模板自动生成验证报告，兼容新旧两种模板格式，支持自动解析用例名称、优先级映射、需求列表和影响范围表填充；当用户需要从测试用例文件批量生成测试报告、更换报告模板、或自动化生成上线后验证报告时使用
dependency:
  python:
    - openpyxl>=3.0.0
    - python-docx>=0.8.11
---

# 测试报告生成器

## 任务目标
- 本 Skill 用于: 从测试用例 xls/xlsx 文件中提取用例数据，按 docx 模板自动生成上线后验证报告
- 能力包含: 解析测试用例文件、克隆模板表格结构、填充用例数据、生成统计总结
- 触发条件: 用户提供测试用例文件和报告模板，需要批量生成格式化的测试验证报告

## 前置准备
- 依赖安装: `pip install openpyxl python-docx`
- 测试用例文件: xls 或 xlsx 格式，需包含以下列(第一行为表头):
  - 用例名称 | 前置条件 | 描述 | 优先级(高/中/低) | 描述步骤 | 预期结果 | 用例类型 | 上线后验证类型
- 报告模板: docx 格式，需包含以下结构:
  - 基本信息表(报告编号/执行人员/日期/时间/公司)
  - Web 测试环境表(浏览器/账号)
  - "验证执行结果" Heading 1 段落
  - "需求：" Heading 2 段落
  - 至少一个测试用例 Heading 3 + 详情表格(8行x5列) + "测试截图"段落
  - "测试总结" Heading 1 段落

## 操作步骤
- 标准流程:
  1. 确认用户提供了测试用例文件和模板文件的路径
  2. 收集可选参数(执行人员、公司、浏览器、账号、时间等)
  3. 调用脚本生成报告:
     ```
     python scripts/generate_report.py \
       --testcase <测试用例文件路径> \
       --template <模板文件路径> \
       --output <输出文件路径> \
       --executor "执行人员" \
       --company "公司名称" \
       --browser "Google Chrome" \
       --account "测试账号" \
       --start-time "21:00" \
       --end-time "22:00"
     ```
  4. 脚本输出 JSON 结果，包含状态、输出文件路径、用例数量、项目信息
  5. 告知用户报告已生成及文件位置

- 可选参数说明:
  - `--report-no`: 报告编号，默认从文件名自动提取
  - `--write-date`: 编写日期，默认当天(格式 YYYY.MM.DD)
  - `--requirement`: 需求描述，默认从文件名提取功能名称
  - `--title`: 报告标题，默认从文件名生成

## 使用示例
- 示例1: 基础用法
  - 场景/输入: 用户提供测试用例 xls 文件和 docx 模板
  - 调用: `python scripts/generate_report.py --testcase ./测试用例.xls --template ./模板.docx --output ./报告.docx --executor "张三" --company "XX公司"`
  - 预期产出: 生成包含所有测试用例的 docx 验证报告
  - 关键要点: 脚本自动从文件名提取项目信息和日期，从用例名称提取验证类型

- 示例2: 完整参数
  - 场景/输入: 用户需要精确控制报告中的所有元数据
  - 调用: `python scripts/generate_report.py --testcase ./用例.xlsx --template assets/template.docx --output ./报告.docx --executor "李四" --company "公司" --browser "Chrome" --account "admin/123" --start-time "20:00" --end-time "22:00" --report-no "报告标题" --requirement "需求：运维费用批量生效失效功能"`
  - 预期产出: 所有元数据按用户指定值填充的报告
  - 关键要点: 所有可选参数均可覆盖自动提取的值

- 示例3: 使用内置模板
  - 场景/输入: 用户没有自己的模板，使用 Skill 内置的标准模板
  - 调用: `python scripts/generate_report.py --testcase ./用例.xls --template assets/template.docx --output ./报告.docx --executor "王五"`
  - 预期产出: 按内置模板格式生成的报告
  - 关键要点: assets/template.docx 是标准上线后验证报告模板

## 资源索引
- 脚本: 见 [scripts/generate_report.py](scripts/generate_report.py)(用途: 解析测试用例文件并生成 docx 报告，参数见操作步骤)
- 资产: 见 [assets/template.docx](assets/template.docx)(用途: 标准上线后验证报告模板，可直接作为 --template 参数使用)

## 注意事项
- 测试用例文件虽然扩展名为 .xls，如果实际是 xlsx 格式(ZIP魔数)，脚本会自动处理
- 用例名称中的【A20260123】【项目组验证】等前缀会被自动解析提取日期和验证类型
- 优先级自动映射: 高->高级，中->中级，低->低级
- 测试步骤和预期结果中的换行符会被保留到报告表格中
- 实际结果默认与预期结果相同(表示测试通过)
- 模板中的测试用例表格结构(含合并单元格)会被完整保留

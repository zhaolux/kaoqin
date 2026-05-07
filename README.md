# 考勤汇总生成

`kaoqin` 用于根据钉钉导出的月度汇总表，生成考勤汇总 Excel。

这是一个标准 `src` 布局的 Python 项目，核心代码位于 [src/kaoqin](</Users/zhaolu/workspace/kaoqin/src/kaoqin>)。

## 安装

开发模式安装：

```bash
pip install -e .
```

安装开发依赖：

```bash
pip install -e '.[dev]'
```

## 运行方式

默认运行：

```bash
python3 generate_attendance.py
python -m kaoqin
kaoqin-generate
```

Excel 源文件和生成文件统一放在 `data/` 目录下。

不传输入文件时，脚本会自动扫描 `data/` 目录下的 `.xlsx` 文件，优先选择文件名包含 `考勤` 且不包含 `自动生成` 的文件。

不传 `-s/--sheet` 时，脚本会优先使用 `月度汇总`，找不到时再自动识别符合钉钉月度汇总格式的工作表。

指定输入文件：

```bash
python -m kaoqin 'data/3月考勤电力，邮政.xlsx'
```

指定工作表和输出文件：

```bash
kaoqin-generate 'data/3月考勤电力，邮政.xlsx' -s '月度汇总' -o 'data/3月考勤结果.xlsx'
```

指定配置目录：

```bash
kaoqin-generate --config-dir ./my-config
```

指定日志级别和日志文件：

```bash
kaoqin-generate --log-level DEBUG --log-file logs/custom.log
```

## 配置文件

脚本默认直接读取包内配置：

- `src/kaoqin/resources/holiday_calendars.json`
- `src/kaoqin/resources/attendance_layout.json`
- `src/kaoqin/resources/attendance_config.json`
- `src/kaoqin/resources/attendance_rules.json`

如果你需要自定义配置，再通过 `--config-dir` 指定一个外部目录覆盖这些默认值。

## 数据目录

`data/` 目录用于存放：

- 钉钉导出的 Excel 源文件
- 脚本自动生成的汇总 Excel 文件

## 日志与异常

- 默认日志文件：`logs/kaoqin.log`
- 默认同时输出到控制台和日志文件
- 可通过 `--log-level` 调整日志级别
- 可通过 `--log-file` 指定日志文件路径
- 输入文件缺失、工作表识别失败、配置格式错误、生成失败等情况会返回明确错误信息并写入日志

### `holiday_calendars.json`

控制节假日和调休日。

字段说明：

- `holidays`: 放假日期列表，格式为 `YYYY-MM-DD`
- `workdays`: 调休上班日期列表，格式为 `YYYY-MM-DD`

### `attendance_layout.json`

控制部门、人员和每个人显示哪些考勤类型行。

字段说明：

- `departments`: 部门列表
- `name`: 部门名或人员名
- `attendance_types`: 该人员在表中要生成的考勤类型行

### `attendance_config.json`

控制一些辅助业务配置。

字段说明：

- `auto_workday_people`: 正常出勤按工作日自动填 `8` 的人员
- `name_alias`: 钉钉姓名和汇总表姓名不一致时的映射

### `attendance_rules.json`

控制状态映射和右侧元数据。

字段说明：

- `status_to_type`: 钉钉状态如何写入到考勤行
- `attendance_type_meta`: 每个考勤类型对应的 `单位` 和 `类型`

## 当前规则

- 按钉钉表头自动识别年份和月份
- 休息日列显示蓝色，工作日无底色
- `合计` 列写入 Excel 公式
- `出勤 √` 行合计会把该人所有 `单位=小时` 的行一起折算为天，按 `8小时=1天`

## 依赖

安装运行依赖：

```bash
pip install -r requirements.txt
```

## 测试

运行全部测试：

```bash
python -m pytest -q
```

仓库已包含 GitHub Actions 工作流：

- push 到 `main`
- 发起 Pull Request

都会自动执行测试。

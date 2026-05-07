# 考勤汇总生成

`kaoqin` 用于根据钉钉导出的月度汇总表，生成考勤汇总 Excel。

这是一个标准 `src` 布局的 Python 项目，核心代码位于 [src/kaoqin](</Users/zhaolu/workspace/kaoqin/src/kaoqin>)。

## 安装

开发模式安装：

```bash
pip install -e .
```

## 运行方式

默认运行：

```bash
python3 generate_attendance.py
python -m kaoqin
kaoqin-generate
```

不传输入文件时，脚本会自动扫描当前目录下的 `.xlsx` 文件，优先选择文件名包含 `考勤` 且不包含 `自动生成` 的文件。

不传 `-s/--sheet` 时，脚本会优先使用 `月度汇总`，找不到时再自动识别符合钉钉月度汇总格式的工作表。

指定输入文件：

```bash
python -m kaoqin '3月考勤电力，邮政.xlsx'
```

指定工作表和输出文件：

```bash
kaoqin-generate '3月考勤电力，邮政.xlsx' -s '月度汇总' -o '3月考勤结果.xlsx'
```

指定配置目录：

```bash
kaoqin-generate --config-dir ./my-config
```

## 配置文件

脚本读取配置时，优先使用当前目录或 `--config-dir` 指定目录下的 JSON；如果找不到，再回退到包内默认配置。

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

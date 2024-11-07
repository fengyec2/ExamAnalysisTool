# 考试成绩分析工具

一个简单的考试成绩分析工具

## 目录结构

```
.
├── ExamAnalysisTool.py  # 主程序文件
└── README.md            # 本文档
```

## 特性

- 生成进退步系数报表
- 生成年级排名折线图
- 生成历次考试成绩单

## 需求

- Python 3.6 及以上版本
- 程序使用 Python 3.10.11 打包，如需在 Windows 7 上使用程序需自行构建
- 依赖库：
  - pandas
  - matplotlib
  - tkinter

## 安装依赖

```bash
pip install pandas matplotlib
```

## 自行构建

```bash
pyinstaller ExamAnalysisTool.spec
```

## 使用说明

1. **运行程序**：在终端中运行以下命令：

   ```bash
   python ExamAnalysisTool.py
   ```

2. **选择文件**：选择一个或多个包含成绩数据的 Excel 文件。文件要求至少包含以下列：
   - `考试编号`
   - `同学`
   - `年级排名`

## 文件格式

导入的 Excel 文件应包含以下示例格式：

| 考试编号 | 同学   | 年级排名 |
|----------|--------|----------|
| 1        | 张三  | 5        |
| 1        | 李四  | 3        |
| 2        | 张三  | 4        |
| 2        | 李四  | 2        |

## 注意事项

- 请确保 Excel 文件的格式正确

## 许可证

本项目采用 GPL v3.0 许可证，详细信息请查看 [LICENSE](LICENSE) 文件

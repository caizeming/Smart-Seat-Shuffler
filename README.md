# 🎓 Smart Seat Shuffler (智能班级排座系统)

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![Version](https://img.shields.io/badge/Version-1.1-blue.svg)
![License](https://img.shields.io/badge/License-CC%20BY--NC%204.0-lightgrey.svg)

一款专为教师打造的现代化、高颜值的班级座位随机分配系统。告别传统手工排座的繁琐，一键生成美观的 Excel 座位表。

## 📥 免安装版（推荐）

普通用户无需安装 Python，直接下载 [**最新 Release**](https://github.com/caizeming/Smart-Seat-Shuffler/releases) 中的 `SmartSeatShuffler-v1.x.x-windows-x64.exe`，双击即用。

> 首次运行如遇 Windows SmartScreen 提示，点击「更多信息 → 仍要运行」即可（exe 未做代码签名的正常现象）。

## ✨ 核心特性 (Features)
- **🎨 现代设计系统**：基于 CustomTkinter 打造，明暗双主题统一配色，主次分明的操作布局。
- **⚡ 高性能**：表格批量刷新、缓存防抖写入、撤销栈上限管理，大数据量也不卡顿。
- **🛡️ 终极数据防线**：无限撤销 + 自动缓存 + 导出自动备份，断电也不怕数据丢失。
- **🧹 智能脏数据预警**：自动检测空姓名、重复学号，表格内双击即可原位修改。
- **📦 灵活规则**：支持紧凑/分散排座，可自定义排除特定损坏座位（如坏电脑）。
- **⌨️ 丝滑交互**：像素级滚动引擎，全局自定义快捷键（Ctrl+Z/S/R）。

## 🚀 快速开始 (Quick Start)
1. 克隆本项目：`git clone https://github.com/caizeming/Smart-Seat-Shuffler.git`
2. 安装依赖：`pip install -r requirements.txt`
3. 运行程序：`python seat_shuffler.py`

## 📋 更新日志 (Changelog)

### v1.1.0
- ⚡ 性能优化：批量表格刷新、缓存防抖合并写入、撤销栈内存上限、座位号安全排序。
- 🎨 界面重构：统一设计配色、修复侧边栏布局问题、操作按钮分级排布、筛选栏表单化对齐、表格列宽自适应。
- 📦 新增 Windows 免安装版下载。

### v1.0.0
- 🎉 首个公开版本发布。

## 📄 声明
本项目仅供学习与教育日常使用，**严禁用于任何商业用途**。

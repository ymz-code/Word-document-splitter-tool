# Word 文档拆分工具 - 文字与图片提取

一款桌面小工具，批量解析 `.docx`，自动把文档正文导出为 TXT，文档内全部图片单独提取保存。

## ✨ 功能特性
- **批量导入**：支持单选文件 / 整个文件夹导入 docx
- **批量提取文字**：每个文档生成独立 txt，保存全部文本内容
- **批量提取图片**：自动收集文档内所有图片，按文档分文件夹存放
- **任务统计**：执行结束弹窗展示成功、失败数量与输出路径
- **运行日志**：实时打印导入、提取信息，方便排查问题
- **日志可视化**：直观查看当前提取进度与报错原因

## 🖼️ 界面预览
### 主界面
<img width="622" height="552" alt="word拆分主页" src="https://github.com/user-attachments/assets/990d0dd5-b882-4916-8d4c-49510833522a" />


### 批量加载文档
<img width="622" height="552" alt="上传" src="https://github.com/user-attachments/assets/c08f912d-d4b1-4986-80de-365f976643d4" />


### 提取完成弹窗
<img width="622" height="552" alt="开始" src="https://github.com/user-attachments/assets/f12c5a10-e4bd-4c4b-91df-f241393b079c" />


## 🚀 使用方法
1. **选择 Word 文件**
   - `选择文件`：手动挑选若干 `.docx`
   - `选择文件夹`：扫描文件夹下全部 docx
   - `清空列表`：清空待处理队列
2. **设置输出目录**
   点击浏览，指定文字、图片的存放根目录
3. **开始提取**
   点击【开始提取】，工具依次处理队列内文档
4. **查看结果**
   - 进度条展示处理进度
   - 日志框实时输出运行信息
   - 处理完毕弹出提示，查看成功/失败统计

## 🚀 安装与启动
### 克隆仓库
```bash
git clone https://github.com/ymz-code/Word-document-splitter-tool.git
cd Word-document-splitter-tool


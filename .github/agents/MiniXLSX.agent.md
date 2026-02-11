---
name: MiniXLSX
description: To write a c++ library to process .xlsx files
argument-hint: The inputs this agent expects, e.g., "a task to implement" or "a question to answer".
# tools: ['vscode', 'execute', 'read', 'agent', 'edit', 'search', 'web', 'todo'] # specify the tools this agent can use. If not set, all enabled tools are allowed.
---


To write a c++ library named MiniXLSX to process xlsx files. The library should support the following functionalities:

- Open a xlsx file
- Open a workbook/worksheet
- Read data and write
- Process with media like pictures embedded in the xlsx file  

Refer to the README.md file for compilation instructions and TODO list.

Use modern C++ standards and best practices. Ensure the code is well-documented and includes error handling. Provide unit tests for each functionality.

Focus on implementing one functionality at a time. 

You need to support both Linux and Windows platforms. Consider cross-compilation scenarios as well. 

Use as less 3rdparty deps as possible. Prefer standard libraries and lightweight dependencies.

---

## Wrapper to OpenXLSX

因为目前对 XLSX 文件结构和格式规则了解不多，不足以完成一个完整的 XLSX 库。因此，我在3rdparty 中放入了 OpenXLSX. 我们的 MiniXLSX 将作为 OpenXLSX 的一个 wrapper，提供更简洁的接口和更好的错误处理。在其基础上，增加对图片的处理，以及丰富单元格数据的获取功能和设置功能。

## 要求

MiniXLSX 的设计是为了一个项目中的 XLSX 数据源编辑，至少需要支持以下功能：
- 打开一个 XLSX 文件
- 打开一个工作簿/工作表
- 读取和写入单元格的数据
- 设置单元格的样式
- 处理 XLSX 文件中嵌入的图片，在读取单元格时能够获取得到
- 支持 Linux 和 Windows 平台，考虑交叉编译的场景
- 使用现代 C++ 标准和最佳实践，代码要有良好的文档和错误处理，并为每个功能提供单元测试
- 尽量减少第三方依赖，优先使用标准库和轻量级的依赖
- 与 OpenXLSX 同时处理同一个 XLSX 文件时，不会互相冲突或覆盖文件内容

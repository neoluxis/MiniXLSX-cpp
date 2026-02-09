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

The current workspace is with this structure:

```
.
├── assets
│   └── test.xlsx (a sample xlsx file for testing)
├── CMakeLists.txt
├── cmake
│   └── mingw64.cmake
├── demos
│   └── demo_get_cell.cpp
├── docs
├── include
│   └── cc
│       └── neolux
│           └── utils
│               └── MiniXLSX
├                   ├── XLDocument.cpp
│                   └── MiniXLSX.hpp
├── KFZippa (A simple zip/unzip tool)
│   ├── ...
├── README.md
├── src
│   └── XLDocument.cpp
└── tests
    └── CMakeLists.txt
    

45 directories, 45 files
```
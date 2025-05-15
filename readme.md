# excel-to-xml-converter

Allow to convert Excel `.xlsx` files into `.xml` files | This example was built to simulate report issuing to the Angola Central Bank

## 🚀 Beginning

The purpose of this project is to create a script that can read an Excel file `XLSX` and displays its contents in `XML` format.

### 🔧 Prerequisites

To develop the project, the following tools must first be installed:

- Code editor: **[Visual Studio Code](https://code.visualstudio.com/)** ![VS Code](/assets/img/vs_code_logo.png "VS Code")
- `Windows` `GCC` Compiler: **[MinGW-w64](https://www.mingw-w64.org/downloads/)** ![MinGW-w64](/assets/img/gcc_compiler.png "MinGW-w64 GCC Compiler")
- Support library (`XLSX` reading): **[libxlsxio](https://sourceforge.net/projects/xlsxio/files/0.2.31/xlsxio-0.2.31-binary-win64.zip/download/)**
- Compiler for support library **`libxlsio`** (optional): **[CMake](https://cmake.org/download/)** ![CMake](/assets/img/cmake.png "CMake")
- VS Code extension to support IntelliSense and build, etc: **[Microsoft C/C++ Extension for VS Code](https://marketplace.visualstudio.com/items?itemName=ms-vscode.cpptools)** ![C/C++ VS Code Extension](/assets/img/vs_code_extension.png "VS Code Extension")
- **[Git](https://git-scm.com/downloads)** ![Git](/assets/img/git_logo.png "Git")

❗ **Important**: After installing MinGW, add the bin path to the system PATH.

```makefile
C:\mingw-w64\bin
```

### ⌨️ Code
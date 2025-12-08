# IDE配置Java 17指南

本项目要求使用 **Java 17或更高版本**。如果IDE配置了Java 8或其他低版本，运行测试时会出现 `UnsupportedClassVersionError` 错误。

## 检查当前Java版本

在终端中运行：
```bash
java -version
```

应该显示类似：
```
openjdk version "17.0.17" 2025-10-21
```

如果显示的是Java 8或更低版本，需要配置IDE使用Java 17。

## IntelliJ IDEA 配置

### 1. 配置项目SDK

1. 打开 **File -> Project Structure** (或按 `Cmd+;` / `Ctrl+Alt+Shift+S`)
2. 在左侧选择 **Project**
3. 在 **SDK** 下拉菜单中选择 **Java 17**
   - 如果没有Java 17，点击下拉菜单旁边的 **New...** 按钮
   - 选择 **JDK**
   - 浏览到Java 17的安装路径（通常在 `/Library/Java/JavaVirtualMachines/` 或 `/usr/libexec/java_home -V` 显示的位置）
4. 在 **Language level** 下拉菜单中选择 **17**
5. 点击 **OK**

### 2. 配置Maven运行器

1. 打开 **File -> Settings** (或按 `Cmd+,` / `Ctrl+Alt+S`)
2. 导航到 **Build, Execution, Deployment -> Build Tools -> Maven -> Runner**
3. 在 **JRE** 下拉菜单中选择 **Java 17**
4. 点击 **OK**

### 3. 配置运行配置

1. 打开 **Run -> Edit Configurations...**
2. 选择要运行的测试类配置（或创建新的配置）
3. 在 **JRE** 下拉菜单中选择 **Java 17**
4. 点击 **OK**

### 4. 验证配置

运行测试类时，控制台应该显示：
```
当前Java版本: 17.0.17
```

而不是：
```
错误：此项目需要Java 17或更高版本，当前版本: 8
```

## Eclipse 配置

### 1. 添加Java 17 JRE

1. 打开 **Window -> Preferences**
2. 导航到 **Java -> Installed JREs**
3. 点击 **Add...**
4. 选择 **Standard VM**，点击 **Next**
5. 点击 **Directory...**，浏览到Java 17的安装目录
6. 点击 **Finish**
7. 勾选新添加的Java 17作为默认JRE
8. 点击 **Apply and Close**

### 2. 配置编译器

1. 打开 **Window -> Preferences**
2. 导航到 **Java -> Compiler**
3. 将 **Compiler compliance level** 设置为 **17**
4. 点击 **Apply and Close**

### 3. 配置项目

1. 右键点击项目，选择 **Properties**
2. 选择 **Java Build Path -> Libraries**
3. 展开 **Modulepath** 或 **Classpath**
4. 如果看到Java 8的JRE，点击 **Remove**
5. 点击 **Add Library...**
6. 选择 **JRE System Library**，点击 **Next**
7. 选择 **Workspace default JRE (Java 17)**，点击 **Finish**
8. 点击 **Apply and Close**

### 4. 配置运行配置

1. 右键点击测试类，选择 **Run As -> Run Configurations...**
2. 选择或创建配置
3. 在 **JRE** 标签页中，选择 **Java 17**
4. 点击 **Run**

## 验证配置成功

运行任何测试类（如 `TestTemplateExtraction`），应该看到：

```
当前Java版本: 17.0.17
=== 模板提取功能单步调试测试 ===
...
```

如果看到错误信息，说明配置未生效，请重新检查上述步骤。

## 常见问题

### Q: 找不到Java 17安装位置

在macOS上，运行：
```bash
/usr/libexec/java_home -V
```

会显示所有已安装的Java版本及其路径。

### Q: 如何安装Java 17

**macOS (使用Homebrew)**:
```bash
brew install openjdk@17
```

**Linux (Ubuntu/Debian)**:
```bash
sudo apt-get install openjdk-17-jdk
```

**Windows**:
从 [Oracle官网](https://www.oracle.com/java/technologies/javase/jdk17-archive-downloads.html) 或 [Adoptium](https://adoptium.net/) 下载安装。

### Q: 配置后仍然使用Java 8

1. 重启IDE
2. 检查环境变量 `JAVA_HOME` 是否指向Java 17
3. 在IDE的终端中运行 `java -version` 确认版本

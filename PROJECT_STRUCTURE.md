# 项目目录结构优化说明

## 优化后的包结构

### 原包结构 → 新包结构

1. **`com.pptfactory.templateengine`** → **`com.pptfactory.template.engine`**
   - 更符合 Java 包命名规范（使用点分隔）
   - 更清晰地表达层级关系

2. **`com.pptfactory.styles`** → **`com.pptfactory.style`**
   - 使用单数形式，更符合 Java 命名习惯
   - 简化包名

3. **`com.pptfactory.lm`** → **`com.pptfactory.ai`**
   - 使用更通用的 `ai` 替代 `lm`（Language Model）
   - 更易理解

4. **`com.pptfactory.app`** → **`com.pptfactory.cli`**
   - 更明确地表示这是命令行接口（CLI）
   - 如果将来有 GUI 版本，可以添加 `gui` 包

## 优化后的目录结构

```
src/main/java/com/pptfactory/
├── template/              # 模板相关
│   └── engine/           # 模板引擎
│       └── PPTTemplateEngine.java
├── style/                # 风格策略（单数形式）
│   ├── StyleStrategy.java
│   ├── DefaultStyle.java
│   ├── ChineseStyle.java
│   ├── MathStyle.java
│   ├── FinanceStyle.java
│   └── SafetyStyle.java
├── ai/                   # AI/大模型相关（替代 lm）
│   ├── ContentGenerator.java
│   └── LayoutClassifier.java
└── cli/                  # 命令行接口（替代 app）
    └── GeneratePPT.java
```

## 优化优势

1. **更规范的包命名**：符合 Java 包命名最佳实践
2. **更清晰的层级**：使用点分隔的包名更易理解
3. **更好的可扩展性**：为未来功能扩展预留空间
4. **更易维护**：包名更直观，降低维护成本

## 迁移步骤

1. 更新所有 Java 文件的 `package` 声明
2. 更新所有 `import` 语句
3. 更新 `pom.xml` 中的主类路径
4. 更新文档中的包名引用


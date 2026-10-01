# 用户指南

## 1. 转换公式

### 当前公式

- 选中包含或不包含定界符的 LaTeX，单击“转换当前公式”。
- 未选中文字时，将转换光标所在或光标前最近的公式。

示例：

```latex
$E=mc^2$
\(x_k=f(x_{k-1})\)
```

### 当前段落

把光标放在段落中，单击“转换当前段落”。同一段落中的多个公式会一次完成转换。

### 全文

单击“转换全文”扫描正文中的所有受支持定界符。转换从后向前进行，以保持 Word Range 位置稳定；某个公式失败不会阻止其他公式。

建议首次对重要论文操作前保存副本。整次命令会写入 Word 的单个撤销记录，可用 `Ctrl+Z` 撤销。

## 2. 自动模式

单击“自动模式”开启。输入完整公式后再输入空格或常见中英文标点，例如：

```latex
$x_k=f(x_{k-1})$ 
```

插件会在短暂延迟后只转换刚完成的公式。自动模式状态会保存；延迟可在“插件设置”中调整。

为了避免打断普通文本，自动模式只处理带受支持定界符的完整公式，不会把任意字母自动改成公式。

## 3. 块公式与矩阵

块公式：

```latex
$$
\frac{a}{b}
$$
```

或：

```latex
\[
\sum_{i=1}^{n}x_i
\]
```

矩阵：

```latex
\begin{bmatrix}
a & b \\
c & d
\end{bmatrix}
```

可用环境包括 `matrix`、`smallmatrix`、`pmatrix`、`bmatrix`、`vmatrix`、`Vmatrix`、`cases`、`array`、`aligned`、`alignedat`、`gathered`、`split`、`align` 和 `equation`（包括带 `*` 的常用形式）。

常用结构示例：

```latex
\dfrac{a}{b}, \binom{n}{k}, \sqrt[3]{x}
\left\langle x,y\right\rangle, \left\|\mathbf{x}\right\|_2
\overline{AB}, \underline{x}, \boxed{x+y}
\overset{!}{=}, \underset{x}{\min}, \underbrace{x_1+\cdots+x_n}_{n\text{ terms}}
\iint_\Omega f\,dA, \oint_C x\,dx
```

`\frac`、`\dfrac`、`\tfrac`、`\cfrac`、`\binom`，可选次方根，常用重音、箭头、关系符、集合符、逻辑符、圆圈运算符和多重积分均可直接转换。

## 4. 科研论文写法

```latex
\mathbf{x}
\hat{\mathbf{x}}
\boldsymbol{\omega}
\operatorname{argmin}_{\mathbf{x}}
\mathrm{diag}(\mathbf{x})
\bm{\theta}
\mathbb{R}^{3\times3}
\mathcal{L}, \mathfrak{g}
SO(3), SE(3), FMCW, LiDAR, IMU
```

`\mathbf`/`\textbf` 映射为 Word 粗体数学样式，`\boldsymbol`/`\bm` 映射为粗斜体，`\mathrm`、`\mathsf`、`\mathtt` 与 `\operatorname` 映射为正体。`\mathbb`、`\mathcal`/`\mathscr` 和 `\mathfrak` 使用 Unicode 数学字母，仍然是可编辑文本而非图片。常见工程缩写自动使用正体运行样式。

## 5. 公式编号

1. 把光标放在包含公式的段落。
2. 单击“公式编号”。
3. 插件为段落设置居中与右对齐制表位，并插入 `(1)` 形式的 `SEQ Equation` 域。

编号是 Word 域，不是普通文本。增删公式后按 `Ctrl+A`、`F9` 更新全文编号。

## 6. 引用公式

1. 把光标放在需要引用的位置。
2. 单击“引用公式”。
3. 在列表中选择目标公式。

插入结果使用 Word `REF` 域并带超链接，例如 `(3)`。编号改变后按 `Ctrl+A`、`F9` 同步引用。

## 7. 失败处理

暂不支持的命令不会被替换为图片，也不会被静默删除。插件会保留原始 LaTeX，并在转换摘要中显示错误。常见不支持内容包括：

- 自定义命令和宏包命令
- `\newcommand`、TikZ、化学公式宏包
- 完整 TeX 排版、宏展开、间距与断行算法
- 复杂多行对齐标签

可把复杂公式拆成受支持的基本结构，或先在 Word 公式编辑器中手动完成不支持部分。

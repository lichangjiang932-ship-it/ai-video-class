# 模块十三：ComfyUI工作流——专业创作者进阶

> **模块概述**：ComfyUI是AI图像和视频生成领域最强大的节点式工作流工具，被专业创作者广泛使用。本模块用8个课时从环境搭建到实战项目，带你掌握ComfyUI的核心工作流，实现角色一致性、ControlNet精准控制、批量生产等专业级能力。

---

## 第91课：ComfyUI环境搭建

### 学习目标

1. 了解ComfyUI的技术架构和核心优势
2. 掌握ComfyUI的安装和配置方法
3. 熟悉ComfyUI的基础界面和操作
4. 理解节点式工作流的核心概念

---

### 91.1 ComfyUI技术概览

#### 91.1.1 什么是ComfyUI

ComfyUI是一个基于节点的**Stable Diffusion图形界面**，允许用户通过连接不同的节点来构建复杂的图像/视频生成工作流。

```
ComfyUI vs 传统WebUI

传统WebUI（如AUTOMATIC1111）：
- 固定界面，参数有限
- 适合简单任务
- 入门门槛低

ComfyUI：
- 节点式工作流，完全可定制
- 适合复杂任务和批量生产
- 入门门槛高，但上限更高
- 支持工作流复用和分享
```

#### 91.1.2 核心优势

| 优势 | 说明 |
|------|------|
| **完全可定制** | 任意组合节点构建工作流 |
| **可视化** | 节点连接清晰可见 |
| **可复用** | 保存和加载工作流模板 |
| **高性能** | 优化的采样和显存管理 |
| **插件生态** | 丰富的第三方插件 |
| **批量生产** | 支持批量处理和自动化 |

---

### 91.2 系统要求

#### 91.2.1 硬件配置

| 配置项 | 最低要求 | 推荐配置 |
|--------|---------|---------|
| **操作系统** | Windows 10 / Linux | Windows 11 / Ubuntu 22.04 |
| **显卡** | NVIDIA 8GB VRAM | NVIDIA 12GB+ VRAM |
| **内存** | 16GB | 32GB |
| **硬盘** | 50GB可用空间 | 100GB+ SSD |
| **Python** | 3.10+ | 3.10.x |

#### 91.2.2 软件依赖

```
必备软件：
- Python 3.10.x
- Git
- CUDA Toolkit（与显卡驱动匹配）
- pip（Python包管理器）

可选软件：
- Visual Studio Code（编辑工作流）
- FFmpeg（视频处理）
- Conda（Python环境管理）
```

---

### 91.3 安装步骤详解

#### 91.3.1 Windows安装

**Step 1：安装Python**

1. 下载Python 3.10.x：https://www.python.org/downloads/
2. 运行安装程序
3. **勾选"Add Python to PATH"**（重要！）
4. 选择"Install Now"
5. 验证安装：打开命令提示符，输入 `python --version`

**Step 2：安装Git**

1. 下载Git：https://git-scm.com/downloads
2. 运行安装程序，使用默认设置
3. 验证安装：输入 `git --version`

**Step 3：下载ComfyUI**

```bash
# 打开命令提示符，进入你想安装的目录
cd C:\Users\你的用户名

# 克隆ComfyUI仓库
git clone https://github.com/comfyanonymous/ComfyUI.git

# 进入ComfyUI目录
cd ComfyUI

# 安装依赖
pip install -r requirements.txt
```

**Step 4：下载基础模型**

```
需要下载的模型：

1. SDXL Base Model（推荐）
   - 下载地址：HuggingFace或CivitAI
   - 放入：models/checkpoints/

2. SD 1.5 Model（可选，兼容更多插件）
   - 放入：models/checkpoints/

3. VAE模型（可选）
   - 放入：models/vae/

4. LoRA模型（角色一致性用）
   - 放入：models/loras/
```

**Step 5：启动ComfyUI**

```bash
# 在ComfyUI目录下运行
python main.py

# 浏览器访问
# http://127.0.0.1:8188
```

#### 91.3.2 Linux安装

```bash
# 安装Python和依赖
sudo apt update
sudo apt install python3.10 python3.10-venv python3-pip git

# 创建虚拟环境（推荐）
python3 -m venv venv
source venv/bin/activate

# 克隆ComfyUI
git clone https://github.com/comfyanonymous/ComfyUI.git
cd ComfyUI

# 安装依赖
pip install -r requirements.txt

# 启动
python main.py
```

---

### 91.4 基础界面认知

#### 91.4.1 界面布局

```
ComfyUI界面

┌─────────────────────────────────────────────┐
│  工具栏（保存/加载/清除/队列）               │
├─────────────────────────────────────────────┤
│                                             │
│         节点画布（工作流编辑区）              │
│                                             │
│   ┌─────┐    ┌─────┐    ┌─────┐           │
│   │节点1│───→│节点2│───→│节点3│           │
│   └─────┘    └─────┘    └─────┘           │
│                                             │
├─────────────────────────────────────────────┤
│  节点搜索栏 / 预览窗口                       │
└─────────────────────────────────────────────┘
```

#### 91.4.2 基础节点

| 节点名称 | 功能 | 输入 | 输出 |
|----------|------|------|------|
| **Load Checkpoint** | 加载大模型 | 模型文件 | MODEL + CLIP + VAE |
| **CLIP Text Encode** | 文本编码 | CLIP + 文本 | CONDITIONING |
| **KSampler** | 采样器 | MODEL + CONDITIONING + LATENT | LATENT |
| **VAE Decode** | 解码图像 | LATENT + VAE | IMAGE |
| **Save Image** | 保存图像 | IMAGE | 文件 |
| **Load Image** | 加载图像 | 文件 | IMAGE |
| **Empty Latent Image** | 创建空白画布 | 尺寸 | LATENT |

#### 91.4.3 节点连接

```
连接规则：

1. 输出端口 → 输入端口
2. 同类型端口才能连接
3. 一个输出可连接多个输入
4. 一个输入只能连接一个输出

颜色编码：
- 紫色：MODEL（模型）
- 粉色：CONDITIONING（条件）
- 黄色：LATENT（潜空间）
- 蓝色：IMAGE（图像）
- 绿色：VAE（编解码器）
- 橙色：CLIP（文本编码器）
```

---

### 91.5 基础文生图工作流

#### 91.5.1 最简工作流

```
[Load Checkpoint]
       ↓
[CLIP Text Encode] ← "a beautiful landscape"
[CLIP Text Encode] ← "blurry, low quality"
       ↓
[Empty Latent Image] ← 1024x1024
       ↓
[KSampler] ← steps:30, cfg:7, sampler:euler_a
       ↓
[VAE Decode]
       ↓
[Save Image]
```

#### 91.5.2 参数说明

| 参数 | 推荐值 | 说明 |
|------|--------|------|
| **Steps** | 25-40 | 采样步数，越多越精细 |
| **CFG** | 7-9 | 提示词遵循度，越高越严格 |
| **Sampler** | euler_a / dpmpp_2m | 采样器类型 |
| **Scheduler** | karras | 噪声调度策略 |
| **Seed** | -1（随机） | 随机种子，-1为随机 |
| **尺寸** | 1024x1024（SDXL） | 输出图像尺寸 |

---

### 91.6 常见问题

**Q1：ComfyUI和AUTOMATIC1111 WebUI有什么区别？**
A：ComfyUI是节点式工作流，更灵活、更强大，但入门门槛更高。WebUI是固定界面，更简单易用。建议先用WebUI入门，再进阶到ComfyUI。

**Q2：显存不够怎么办？**
A：(1) 使用--lowvram参数启动；(2) 降低图像尺寸；(3) 减少采样步数；(4) 使用更小的模型。

**Q3：安装失败怎么办？**
A：(1) 检查Python版本是否正确；(2) 检查CUDA是否安装；(3) 检查网络连接；(4) 尝试使用虚拟环境。

**Q4：如何更新ComfyUI？**
A：在ComfyUI目录下执行 `git pull`，然后重启。

---

### 91.7 课后练习

1. **环境搭建**：完成ComfyUI的安装和配置
2. **基础工作流**：搭建并运行最简文生图工作流
3. **参数实验**：调整Steps、CFG、Seed等参数，观察效果变化
4. **模型测试**：下载并测试不同的基础模型

---

## 第92课：ComfyUI核心插件安装

### 学习目标

1. 了解ComfyUI插件生态系统
2. 掌握核心插件的安装方法
3. 理解各插件的功能和用途
4. 学会管理插件的安装和更新

---

### 92.1 ComfyUI Manager（插件管理器）

#### 92.1.1 安装ComfyUI Manager

```bash
cd ComfyUI/custom_nodes
git clone https://github.com/ltdrdata/ComfyUI-Manager.git
cd ComfyUI-Manager
pip install -r requirements.txt
# 重启ComfyUI
```

#### 92.1.2 使用ComfyUI Manager

1. 启动ComfyUI后，在界面中找到"Manager"按钮
2. 点击"Install Custom Nodes"
3. 搜索想要的插件
4. 点击"Install"
5. 重启ComfyUI

---

### 92.2 必装插件详解

#### 92.2.1 IP-Adapter Plus

```
功能：图像特征注入，实现角色一致性
安装：通过ComfyUI Manager搜索"IPAdapter_plus"
依赖模型：ip-adapter-faceid-plusv2_sdxl

核心节点：
- Load IPAdapter Model：加载IP-Adapter模型
- IPAdapter FaceID：面部特征注入
- IPAdapter Advanced：高级设置
```

#### 92.2.2 ControlNet

```
功能：骨骼/边缘/深度等精准控制
安装：通过ComfyUI Manager搜索"ControlNet"
依赖模型：control_v11p_sd15_openpose等

核心节点：
- Load ControlNet Model：加载ControlNet模型
- Apply ControlNet：应用ControlNet
- ControlNet Preprocessor：预处理参考图
```

#### 92.2.3 ADetailer

```
功能：面部检测和自动修复
安装：通过ComfyUI Manager搜索"ADetailer"
依赖模型：face_yolov8m.pt

核心节点：
- ADetailer：面部检测和重渲染
- 面部检测阈值：0.3
- 重渲染强度：0.4-0.6
```

#### 92.2.4 FaceRestore

```
功能：面部重建和修复
安装：通过ComfyUI Manager搜索"FaceRestore"
依赖模型：GFPGAN或CodeFormer

核心节点：
- FaceRestore：面部重建
- 重建强度：0.5-0.8
```

#### 92.2.5 Inspire Pack

```
功能：工具集，包含多种实用节点
安装：通过ComfyUI Manager搜索"Inspire"

核心节点：
- Prompt Schedule：提示词调度
- Image Batch：图像批处理
- Regional Conditioning：区域条件
```

---

### 92.3 插件安装方式

#### 92.3.1 方式一：ComfyUI Manager（推荐）

```
优点：简单快捷，自动处理依赖
缺点：部分插件可能不在列表中

步骤：
1. 打开ComfyUI Manager
2. 点击"Install Custom Nodes"
3. 搜索插件名称
4. 点击"Install"
5. 重启ComfyUI
```

#### 92.3.2 方式二：手动安装

```
优点：可以安装任何插件
缺点：需要手动处理依赖

步骤：
1. 打开命令提示符
2. 进入ComfyUI/custom_nodes目录
3. git clone插件仓库
4. 进入插件目录
5. pip install -r requirements.txt
6. 重启ComfyUI
```

#### 92.3.3 方式三：下载ZIP安装

```
步骤：
1. 从GitHub下载插件ZIP文件
2. 解压到ComfyUI/custom_nodes目录
3. 安装依赖（如果有requirements.txt）
4. 重启ComfyUI
```

---

### 92.4 插件管理

#### 92.4.1 更新插件

```
通过ComfyUI Manager：
1. 打开Manager
2. 点击"Update All"或单独更新
3. 重启ComfyUI

手动更新：
cd ComfyUI/custom_nodes/插件目录
git pull
pip install -r requirements.txt
# 重启ComfyUI
```

#### 92.4.2 禁用/卸载插件

```
禁用插件：
- 在插件目录下创建.disabled文件
- 或在ComfyUI Manager中禁用

卸载插件：
- 删除插件目录
- 或通过ComfyUI Manager卸载
```

---

### 92.5 常见问题

**Q1：插件安装后不显示节点？**
A：(1) 确认已重启ComfyUI；(2) 检查插件是否安装成功；(3) 检查依赖是否安装完整。

**Q2：插件之间冲突怎么办？**
A：(1) 逐个禁用插件排查；(2) 检查插件版本兼容性；(3) 查看ComfyUI控制台错误信息。

**Q3：如何找到需要的插件？**
A：(1) 使用ComfyUI Manager搜索；(2) 在GitHub搜索"ComfyUI"相关插件；(3) 参考社区推荐的插件列表。

---

### 92.6 课后练习

1. **安装Manager**：安装并配置ComfyUI Manager
2. **安装核心插件**：安装IP-Adapter、ControlNet、ADetailer三个核心插件
3. **验证安装**：确认所有插件节点正常显示
4. **插件测试**：使用一个插件节点测试功能

---

## 第93课：ComfyUI文生图工作流

### 学习目标

1. 掌握ComfyUI文生图的完整工作流搭建
2. 理解各节点参数的含义和调优方法
3. 学会使用正向和负向提示词
4. 掌握批量生成的方法

---

### 93.1 基础文生图工作流

#### 93.1.1 完整节点连接

```
[Load Checkpoint] ─── MODEL ──→ [KSampler]
                    CLIP ──→ [CLIP Text Encode+] ──→ [KSampler]
                    CLIP ──→ [CLIP Text Encode-] ──→ [KSampler]
                    VAE ──→ [VAE Decode]

[Empty Latent Image] ──→ [KSampler]

[KSampler] ──→ [VAE Decode]
[VAE Decode] ──→ [Save Image]
```

#### 93.1.2 提示词编写

**正向提示词（想要的内容）：**
```
a beautiful anime girl, long silver hair, blue eyes,
white dress, standing in a cherry blossom garden,
soft sunlight, detailed face, masterpiece, best quality,
1080p, cinematic lighting
```

**负向提示词（不想要的内容）：**
```
blurry, low quality, deformed, ugly, duplicate,
extra limbs, bad anatomy, bad hands, missing fingers,
watermark, text, signature
```

#### 93.1.3 提示词技巧

| 技巧 | 说明 | 示例 |
|------|------|------|
| **权重强调** | 使用(词:权重)强调特定元素 | (beautiful:1.2), (detailed:1.1) |
| **顺序优先** | 越靠前的词权重越高 | 先写主要特征，再写次要特征 |
| **否定排除** | 使用负向提示词排除不需要的内容 | 负向提示词中写"blurry, deformed" |
| **风格描述** | 指定艺术风格 | anime style, realistic, oil painting |
| **质量描述** | 指定画质要求 | masterpiece, best quality, 8k |

---

### 93.2 参数调优指南

#### 93.2.1 采样步数（Steps）

| 步数 | 效果 | 速度 | 推荐场景 |
|------|------|------|---------|
| 15-20 | 快速预览 | 快 | 草稿测试 |
| 25-30 | 平衡 | 中 | **日常使用** |
| 35-40 | 精细 | 慢 | 最终输出 |
| 50+ | 极精细 | 很慢 | 特殊需求 |

#### 93.2.2 CFG Scale

| CFG值 | 效果 | 适用场景 |
|--------|------|---------|
| 3-5 | 自由发挥，创意性强 | 创意探索 |
| 7-9 | 平衡，遵循提示词但不僵化 | **推荐默认** |
| 10-12 | 严格遵循提示词 | 精准控制 |
| 15+ | 过度遵循，可能过饱和 | 不推荐 |

#### 93.2.3 采样器选择

| 采样器 | 特点 | 推荐场景 |
|--------|------|---------|
| **euler_a** | 快速，效果好 | **通用推荐** |
| **dpmpp_2m** | 稳定，质量高 | 高质量输出 |
| **dpmpp_sde** | 细节丰富 | 需要细节的场景 |
| **ddim** | 确定性强 | 需要可复现的结果 |

---

### 93.3 批量生成

#### 93.3.1 批量设置

```
方法1：修改Batch Size
- Empty Latent Image → Batch Size: 4
- 一次生成4张图

方法2：使用Queue
- 修改Seed后多次Queue
- 每次生成不同结果

方法3：使用循环节点
- 安装循环插件
- 自动遍历多个提示词
```

#### 93.3.2 批量筛选

```
筛选流程：
1. 生成4-8张候选图
2. 快速浏览所有结果
3. 标记最佳版本
4. 对最佳版本微调参数
5. 生成最终版本
```

---

### 93.4 常见问题

**Q1：生成的图片和提示词不符？**
A：(1) 降低CFG值；(2) 简化提示词；(3) 调整提示词顺序；(4) 增加采样步数。

**Q2：生成的图片质量差？**
A：(1) 增加采样步数；(2) 使用质量提示词；(3) 检查模型质量；(4) 尝试不同采样器。

**Q3：生成速度太慢？**
A：(1) 减少采样步数；(2) 降低图像尺寸；(3) 使用--lowvram参数；(4) 升级显卡。

---

### 93.5 课后练习

1. **搭建工作流**：搭建完整的文生图工作流
2. **提示词实验**：使用不同提示词生成10张图
3. **参数调优**：测试不同Steps和CFG值的效果
4. **批量生成**：一次生成4张图并筛选最佳

---

## 第94课：ComfyUI图生图工作流

### 学习目标

1. 掌握ComfyUI图生图的工作流搭建
2. 理解降噪强度对生成效果的影响
3. 学会使用局部重绘（Inpainting）功能
4. 掌握基于参考图的风格迁移

---

### 94.1 基础图生图工作流

#### 94.1.1 节点连接

```
[Load Image] ──→ [VAE Encode] ──→ [KSampler]
[Load Checkpoint] ──→ [KSampler]
[CLIP Text Encode+] ──→ [KSampler]
[CLIP Text Encode-] ──→ [KSampler]
[KSampler] ──→ [VAE Decode] ──→ [Save Image]
```

#### 94.1.2 降噪强度

| 降噪值 | 效果 | 适用场景 |
|--------|------|---------|
| 0.3-0.4 | 微调，保留原图大部分内容 | 轻微修改 |
| 0.5-0.6 | 中等修改，保留构图 | **推荐默认** |
| 0.7-0.8 | 较大修改，保留基本特征 | 风格迁移 |
| 0.9-1.0 | 接近重新生成 | 大幅修改 |

---

### 94.2 局部重绘（Inpainting）

#### 94.2.1 节点连接

```
[Load Image] ──→ [VAE Encode (Inpaint)]
[Load Mask] ──→ [VAE Encode (Inpaint)]
[VAE Encode (Inpaint)] ──→ [KSampler]
[KSampler] ──→ [VAE Decode] ──→ [Save Image]
```

#### 94.2.2 操作步骤

1. 加载原始图片
2. 创建或加载蒙版（Mask）
   - 蒙版白色区域 = 需要重绘的区域
   - 蒙版黑色区域 = 保持不变的区域
3. 编写重绘提示词
4. 设置降噪强度（0.6-0.8）
5. 生成

#### 94.2.3 蒙版创建方法

```
方法1：使用Load Mask节点
- 上传预先制作的蒙版图片

方法2：使用绘图节点
- 安装绘图插件
- 在ComfyUI中直接绘制蒙版

方法3：使用外部工具
- 在Photoshop/画图中制作蒙版
- 保存为黑白图片
- 上传到ComfyUI
```

---

### 94.3 风格迁移

#### 94.3.1 基于参考图的风格迁移

```
[Load Checkpoint] ← 风格模型（如动漫风格）
[Load Image] ← 原始图片
[VAE Encode] ← 编码原始图片
[CLIP Text Encode] ← "anime style, vibrant colors"
[KSampler] ← 降噪: 0.6-0.75
[VAE Decode]
[Save Image]
```

#### 94.3.2 风格迁移技巧

| 技巧 | 说明 |
|------|------|
| 选择合适的模型 | 不同模型擅长不同风格 |
| 调整降噪强度 | 太低保留原风格，太高丢失原图 |
| 使用风格提示词 | 在提示词中明确指定目标风格 |
| 保持构图 | 降低降噪强度保持原图构图 |

---

### 94.4 常见问题

**Q1：重绘区域和原图不协调？**
A：(1) 增加降噪强度；(2) 调整重绘提示词；(3) 使用ADetailer修复过渡区域。

**Q2：蒙版边缘不自然？**
A：(1) 使用模糊蒙版边缘；(2) 增加蒙版的渐变区域；(3) 在后期中手动修复。

**Q3：风格迁移效果不明显？**
A：(1) 降低降噪强度；(2) 使用更强的风格提示词；(3) 选择更风格化的模型。

---

### 94.5 课后练习

1. **图生图工作流**：搭建完整的图生图工作流
2. **降噪实验**：测试不同降噪值的效果
3. **局部重绘**：对一张图片进行局部重绘
4. **风格迁移**：将一张照片转换为动漫风格

---

## 第95课：ComfyUI ControlNet深度应用

### 学习目标

1. 理解ControlNet的工作原理
2. 掌握OpenPose、Canny、Depth等ControlNet的使用
3. 学会使用ControlNet精准控制画面
4. 掌握多ControlNet叠加使用

---

### 95.1 ControlNet类型详解

| ControlNet类型 | 功能 | 适用场景 |
|----------------|------|---------|
| **OpenPose** | 骨骼姿态控制 | 人物动作控制 |
| **Canny** | 边缘线条控制 | 建筑、物体轮廓 |
| **Depth** | 深度图控制 | 空间关系、前后景 |
| **Scribble** | 手绘线稿控制 | 创意草图 |
| **Lineart** | 线稿控制 | 线稿上色 |
| **Normal Map** | 法线图控制 | 材质和光影 |
| **SoftEdge** | 柔边控制 | 柔和轮廓 |
| **Shuffle** | 内容重排 | 风格迁移 |

---

### 95.2 OpenPose骨骼控制

#### 95.2.1 操作步骤

```
Step 1：加载OpenPose预处理器
[Load Image] → [OpenPose Preprocessor] → 骨骼图

Step 2：加载ControlNet模型
[Load ControlNet Model] → control_v11p_sd15_openpose

Step 3：应用ControlNet
[Apply ControlNet] ← 骨骼图 + ControlNet模型
强度：0.5-0.8

Step 4：连接到KSampler
[Apply ControlNet] → CONDITIONING → [KSampler]
```

#### 95.2.2 参数设置

| 参数 | 推荐值 | 说明 |
|------|--------|------|
| 强度 | 0.5-0.8 | 控制骨骼姿态的影响程度 |
| 预处理器 | dw_openpose_full | 检测全身骨骼 |
| 起始步 | 0 | 从开始应用 |
| 结束步 | 1 | 到结束应用 |

---

### 95.3 Canny边缘控制

#### 95.3.1 操作步骤

```
Step 1：提取边缘图
[Load Image] → [Canny Preprocessor] → 边缘图
参数：low_threshold: 100, high_threshold: 200

Step 2：加载ControlNet模型
[Load ControlNet Model] → control_v11p_sd15_canny

Step 3：应用ControlNet
[Apply ControlNet] ← 边缘图 + ControlNet模型
强度：0.5-0.7
```

#### 95.3.2 Canny参数调优

| 参数 | 效果 |
|------|------|
| low_threshold低 | 检测更多边缘 |
| low_threshold高 | 只检测强边缘 |
| high_threshold低 | 边缘更细 |
| high_threshold高 | 边缘更粗 |

---

### 95.4 Depth深度控制

#### 95.4.1 操作步骤

```
Step 1：提取深度图
[Load Image] → [Depth Preprocessor] → 深度图

Step 2：加载ControlNet模型
[Load ControlNet Model] → control_v11f1p_sd15_depth

Step 3：应用ControlNet
[Apply ControlNet] ← 深度图 + ControlNet模型
强度：0.5-0.7
```

---

### 95.5 多ControlNet叠加

#### 95.5.1 叠加方法

```
[Apply ControlNet 1] ← OpenPose（骨骼）
       ↓
[Apply ControlNet 2] ← Canny（边缘）
       ↓
[KSampler]
```

#### 95.5.2 叠加建议

| 组合 | 效果 | 推荐度 |
|------|------|--------|
| OpenPose + Canny | 姿态+轮廓 | ★★★★★ |
| OpenPose + Depth | 姿态+空间 | ★★★★☆ |
| Canny + Depth | 轮廓+空间 | ★★★★☆ |
| 三个叠加 | 全面控制 | ★★★☆☆ |

---

### 95.6 常见问题

**Q1：ControlNet效果不明显？**
A：(1) 检查强度设置；(2) 确认预处理器正确；(3) 检查ControlNet模型是否匹配。

**Q2：骨骼检测不准确？**
A：(1) 使用更清晰的参考图；(2) 尝试不同的预处理器；(3) 手动调整骨骼图。

**Q3：多个ControlNet冲突？**
A：(1) 降低各ControlNet的强度；(2) 调整起始步和结束步；(3) 减少ControlNet数量。

---

### 95.7 课后练习

1. **OpenPose测试**：使用OpenPose控制人物姿态
2. **Canny测试**：使用Canny控制建筑轮廓
3. **Depth测试**：使用Depth控制空间关系
4. **多ControlNet叠加**：同时使用OpenPose和Canny

---

## 第96课：ComfyUI角色一致性工作流

### 学习目标

1. 掌握IP-Adapter + LoRA + ADetailer的黄金组合
2. 学会搭建完整的角色一致性工作流
3. 理解各参数对角色一致性的影响
4. 掌握批量生成角色一致图片的方法

---

### 96.1 角色一致性技术栈

```
技术栈

IP-Adapter FaceID → 面部特征注入
LoRA → 角色专属模型
ControlNet OpenPose → 姿态控制
ADetailer → 面部修复
FaceRestore → 面部重建
```

---

### 96.2 完整工作流搭建

#### 96.2.1 节点连接

```
[Load Checkpoint]
       ↓
[Load LoRA] ← 角色专属LoRA模型
       ↓
[Load IPAdapter Model] ← ip-adapter-faceid-plusv2
[Load Image] ← 角色参考图
[IPAdapter FaceID] ← 模型 + 参考图
权重：0.7-0.85
       ↓
[Load ControlNet Model] ← openpose
[Apply ControlNet] ← 骨骼图 + ControlNet
强度：0.5-0.7
       ↓
[CLIP Text Encode] ← 正向提示词
[CLIP Text Encode] ← 负向提示词
       ↓
[Empty Latent Image] ← 目标尺寸
       ↓
[KSampler] ← steps:30, cfg:7
       ↓
[ADetailer] ← 面部检测和修复
阈值：0.3，重渲染强度：0.5
       ↓
[VAE Decode]
       ↓
[Save Image]
```

#### 96.2.2 参数设置

| 参数 | 推荐值 | 说明 |
|------|--------|------|
| IP-Adapter权重 | 0.7-0.85 | 太高过拟合，太低不像 |
| LoRA权重 | 0.6-0.8 | 控制角色特征强度 |
| ControlNet强度 | 0.5-0.7 | 控制姿态但不锁死细节 |
| ADetailer阈值 | 0.3 | 面部检测灵敏度 |
| ADetailer重渲染 | 0.4-0.6 | 面部修复强度 |
| Steps | 30-40 | 采样质量 |
| CFG | 7-9 | 提示词遵循度 |

---

### 96.3 LoRA训练（简要）

#### 96.3.1 训练数据准备

```
数据集要求：
- 数量：30-50张
- 分辨率：512x512 或 1024x1024
- 角度：正面10-15张，侧面5-8张，3/4角度5-8张
- 表情：平静、微笑、愤怒、惊讶等
- 光照：均匀，避免强烈阴影
- 背景：简洁，纯色最佳
```

#### 96.3.2 训练参数

| 参数 | 推荐值 |
|------|--------|
| network_dim | 32-64 |
| learning_rate | 1e-4 ~ 5e-5 |
| max_epochs | 8-15 |
| batch_size | 1-2 |
| resolution | 1024 |

---

### 96.4 批量生成

#### 96.4.1 批量工作流

```
1. 准备多个场景的提示词
2. 准备多个姿态的骨骼图
3. 使用同一个角色参考图和LoRA
4. 批量生成不同场景的角色图片
5. 筛选最佳版本
```

#### 96.4.2 一致性检查

```
检查清单：
- [ ] 面部特征一致（五官位置、形状）
- [ ] 发型发色一致
- [ ] 服装一致
- [ ] 体型一致
- [ ] 整体风格一致
```

---

### 96.5 常见问题

**Q1：角色面部变形怎么办？**
A：(1) 降低IP-Adapter权重；(2) 使用ADetailer修复；(3) 检查参考图质量。

**Q2：角色在不同场景中差异大？**
A：(1) 提高LoRA权重；(2) 使用更高质量的参考图；(3) 增加IP-Adapter权重。

**Q3：工作流运行很慢？**
A：(1) 减少采样步数；(2) 降低图像尺寸；(3) 使用--lowvram参数。

---

### 96.6 课后练习

1. **工作流搭建**：搭建完整的角色一致性工作流
2. **一致性测试**：生成同一角色在5个不同场景的图片
3. **参数调优**：测试不同IP-Adapter和LoRA权重的效果
4. **批量生成**：一次性生成10张角色一致的图片

---

## 第97课：Z-Comics Workflow 2.1

### 学习目标

1. 了解Z-Comics插件的功能和用途
2. 掌握多格漫画的批量生成方法
3. 学会自定义漫画分格布局
4. 掌握漫画风格的统一控制

---

### 97.1 Z-Comics简介

Z-Comics是一个ComfyUI插件，专门用于**批量生成多格漫画**。它支持自定义分格布局、角色一致性控制、风格统一等功能。

```
Z-Comics核心功能

多格布局 → 支持2/3/4/6/9格布局
角色一致性 → 跨格保持角色一致
风格统一 → 所有格子风格一致
批量生成 → 一次生成整页漫画
文字气泡 → 支持添加对话气泡
```

---

### 97.2 安装与配置

```bash
# 通过ComfyUI Manager安装
# 搜索"Z-Comics"并安装

# 或手动安装
cd ComfyUI/custom_nodes
git clone https://github.com/Z-Comics/ComfyUI-Z-Comics.git
pip install -r requirements.txt
```

---

### 97.3 操作步骤

#### Step 1：选择分格布局

```
布局选项：
- 2格：横排/竖排
- 3格：横排/竖排/L型
- 4格：田字格/横排/竖排
- 6格：2x3/3x2
- 9格：3x3
```

#### Step 2：设置每格内容

```
为每个格子设置：
- 提示词（描述画面内容）
- 参考图（角色/场景参考）
- 景别（特写/中景/全景）
- 镜头角度（平视/俯视/仰视）
```

#### Step 3：设置角色参考

```
角色一致性设置：
- 上传角色参考图
- 设置IP-Adapter权重
- 选择LoRA模型（如有）
- 设置ControlNet（如需要）
```

#### Step 4：批量生成

```
生成流程：
1. 点击"Generate All"
2. AI为每个格子生成图片
3. 自动拼接为整页漫画
4. 预览和调整
```

---

### 97.4 漫画风格控制

#### 97.4.1 风格提示词

```
统一风格提示词：
"manga style, black and white, lineart, 
detailed shading, professional comic art"

彩色漫画风格：
"colorful manga style, vibrant colors, 
detailed shading, professional comic art"
```

#### 97.4.2 风格一致性

```
确保风格一致的方法：
1. 所有格子使用相同的基础模型
2. 所有格子使用相同的风格提示词
3. 使用相同的采样器和参数
4. 使用风格LoRA（如有）
```

---

### 97.5 常见问题

**Q1：不同格子的角色不一致？**
A：(1) 使用相同的IP-Adapter参考图；(2) 提高IP-Adapter权重；(3) 使用角色LoRA。

**Q2：格子之间的风格不统一？**
A：(1) 使用相同的风格提示词；(2) 使用相同的模型和参数；(3) 在后期中统一调色。

**Q3：生成的漫画质量不高？**
A：(1) 增加采样步数；(2) 使用更高质量的模型；(3) 优化提示词。

---

### 97.6 课后练习

1. **安装Z-Comics**：安装并配置Z-Comics插件
2. **基础漫画**：生成一页4格漫画
3. **角色一致性**：确保4个格子中角色一致
4. **风格实验**：尝试不同的漫画风格

---

## 第98课：实战项目——搭建完整工作流

### 学习目标

1. 掌握从零搭建完整ComfyUI工作流的方法
2. 学会整合多种技术实现复杂需求
3. 完成一个可用于批量生产的工作流
4. 掌握工作流的保存和复用

---

### 98.1 项目目标

搭建一个完整的漫剧图像生成工作流，包含：
- 角色一致性（IP-Adapter + LoRA）
- 多场景生成（ControlNet）
- 面部修复（ADetailer）
- 批量生产支持
- 工作流模板保存

---

### 98.2 搭建流程

#### Step 1：准备资产

```
需要准备的资产：
- 角色参考图（3-5张）
- 角色LoRA模型（如有）
- 场景参考图（可选）
- ControlNet预处理图（可选）
- 基础模型（SDXL推荐）
```

#### Step 2：搭建基础工作流

```
基础文生图工作流：
[Load Checkpoint] → [CLIP Text Encode] → [KSampler] → [VAE Decode] → [Save Image]
```

#### Step 3：添加IP-Adapter

```
添加IP-Adapter节点：
[Load IPAdapter Model] → [IPAdapter FaceID]
连接角色参考图
设置权重：0.7-0.85
```

#### Step 4：添加LoRA

```
添加LoRA节点：
[Load LoRA] → 连接到Checkpoint后
设置权重：0.6-0.8
```

#### Step 5：添加ControlNet

```
添加ControlNet节点：
[Load ControlNet Model] → [Apply ControlNet]
连接骨骼图/边缘图
设置强度：0.5-0.7
```

#### Step 6：添加ADetailer

```
添加ADetailer节点：
[ADetailer] → 连接到KSampler后
设置阈值：0.3
设置重渲染强度：0.5
```

#### Step 7：测试和优化

```
测试10个不同场景：
1. 室内场景 × 3
2. 室外场景 × 3
3. 特殊场景 × 4

检查清单：
- [ ] 角色一致性
- [ ] 画面质量
- [ ] 场景适配性
- [ ] 生成速度
```

#### Step 8：保存工作流

```
保存方法：
1. 点击"Save"按钮
2. 输入工作流名称
3. 保存为.json文件
4. 可随时加载使用
```

---

### 98.3 工作流优化

#### 98.3.1 性能优化

| 优化方法 | 效果 |
|----------|------|
| 减少采样步数 | 加快速度 |
| 降低图像尺寸 | 减少显存 |
| 使用--lowvram | 低显存优化 |
| 使用高效采样器 | 加快速度 |

#### 98.3.2 质量优化

| 优化方法 | 效果 |
|----------|------|
| 增加采样步数 | 提高质量 |
| 使用高质量模型 | 提高画质 |
| 优化提示词 | 提高准确性 |
| 调整ControlNet强度 | 提高控制精度 |

---

### 98.4 交付物

```
项目交付物：
- 完整的ComfyUI工作流文件（.json）
- 角色参考图集
- LoRA模型文件（如有）
- 工作流使用说明
- 10张测试生成的角色一致图片
```

---

### 98.5 常见问题

**Q1：工作流节点太多，运行很慢？**
A：(1) 优化节点连接，减少冗余；(2) 降低采样步数；(3) 使用更高效的节点。

**Q2：如何分享工作流？**
A：(1) 导出为.json文件；(2) 上传到GitHub或社区；(3) 附上使用说明。

**Q3：工作流在不同电脑上效果不同？**
A：(1) 确保使用相同的模型版本；(2) 确保使用相同的随机种子；(3) 检查显卡差异。

---

### 98.6 课后练习

1. **完整搭建**：从零搭建包含所有组件的完整工作流
2. **测试验证**：使用10个不同场景测试工作流
3. **参数优化**：优化各参数达到最佳效果
4. **保存复用**：保存工作流并尝试在不同项目中复用
5. **分享交流**：将你的工作流分享到社区

---

## 模块总结

### 核心知识点回顾

| 课时 | 核心内容 | 关键技能 |
|------|---------|---------|
| 第91课 | ComfyUI环境搭建 | 安装配置、基础界面、节点认知 |
| 第92课 | 核心插件安装 | Manager、IP-Adapter、ControlNet、ADetailer |
| 第93课 | 文生图工作流 | 提示词、参数调优、批量生成 |
| 第94课 | 图生图工作流 | 降噪强度、局部重绘、风格迁移 |
| 第95课 | ControlNet应用 | OpenPose、Canny、Depth、多ControlNet叠加 |
| 第96课 | 角色一致性工作流 | IP-Adapter+LoRA+ADetailer黄金组合 |
| 第97课 | Z-Comics漫画 | 多格漫画、批量生成、风格统一 |
| 第98课 | 实战项目 | 完整工作流搭建、测试、优化 |

### ComfyUI核心工作流速查

| 工作流 | 核心节点 | 适用场景 |
|--------|---------|---------|
| 文生图 | Checkpoint + CLIP + KSampler | 通用图像生成 |
| 图生图 | Load Image + VAE Encode + KSampler | 基于参考图生成 |
| 角色一致性 | IP-Adapter + LoRA + ADetailer | 角色一致图片 |
| 姿态控制 | OpenPose + ControlNet | 人物动作控制 |
| 边缘控制 | Canny + ControlNet | 建筑/物体控制 |
| 深度控制 | Depth + ControlNet | 空间关系控制 |
| 局部重绘 | Inpaint + Mask + KSampler | 局部修改 |
| 漫画生成 | Z-Comics + 批量节点 | 多格漫画 |

### 进阶学习方向

1. **ComfyUI视频工作流**：将图像工作流扩展到视频生成
2. **自定义节点开发**：学习Python开发自定义ComfyUI节点
3. **API集成**：通过API实现ComfyUI的自动化调用
4. **云端部署**：将ComfyUI部署到云端GPU服务器

---

> **恭喜你完成模块十三的全部学习！** 你已经掌握了ComfyUI的核心工作流，可以开始专业级的AI图像和视频制作了。下一模块我们将学习工业化生产与团队协作，了解如何将个人创作扩展为团队规模的生产能力。

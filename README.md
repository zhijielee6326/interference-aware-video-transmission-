# 面向信号干扰检测与识别的视频传输系统

基于 CCNN 的 7 类干扰信号识别 & 自适应 QPSK 视频传输系统

## 项目逻辑

```
阶段1 数据生成 → 阶段2 模型训练 → 阶段3 系统集成 → 阶段4 测试验证
   ↓                ↓                ↓                ↓
 MATLAB/Python     CCNN训练         视频传输系统       USRP实机验证
 信号生成脚本       模型权重         主系统+检测器      JSR扫描+自适应测试
```

## 目录结构

```
├── 01_数据生成/           # 阶段1：生成训练用干扰信号数据
│   ├── MATLAB信号生成/    # MATLAB版信号生成脚本（旧版，保留对比）
│   ├── datasets/          # 训练/测试数据集
│   └── *.py               # Python版信号生成器
│
├── 02_模型训练/           # 阶段2：CCNN干扰识别模型训练
│   ├── models/            # 模型权重 (best: 99.91%)
│   ├── scripts/           # 训练/推理代码
│   └── results/           # 训练日志与结果
│
├── 03_视频传输系统/       # 阶段3：完整的自适应视频传输系统
│   ├── video_transmission_system.py   # 主系统 (核心)
│   ├── cnn_interference_detector.py   # CCNN检测器
│   └── run_*.sh           # 启动脚本
│
├── 04_硬件测试验证/       # 阶段4：USRP实机测试与实验结果
│   ├── usrp_*.py/json     # 实机JSR测试
│   ├── test_*.py          # 离线测试脚本
│   └── figure_*.png       # 论文实验图表
│
├── docs/                  # 论文文档
│   ├── 毕业论文.md         # 论文 Markdown 源稿
│   ├── 毕业论文_李智杰.docx # 论文 Word 版
│   ├── 答辩PPT_李智杰.pptx # 答辩 PPT
│   └── 论文图表/           # 论文配图
│
├── 进展/                  # 开题报告、任务书、进展报告
├── 参考文献/              # 参考文献PDF/DOCX
└── venv/                  # Python 虚拟环境
```

## 技术栈

| 领域 | 技术 |
|------|------|
| 深度学习 | PyTorch, CNN, SE-Attention, Multi-Head Attention |
| 软件无线电 | UHD / USRP B210, QPSK 调制解调 |
| 视频处理 | OpenCV, JPEG 编解码, 自适应质量控制 |
| 信道编码 | XOR 前向纠错, CRC-16 校验 |

## 项目信息

- 作者：李智杰
- 类型：本科毕业设计
- 指导教师：刘磊
- 学校：北京信息科技大学

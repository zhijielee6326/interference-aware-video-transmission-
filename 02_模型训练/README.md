# 阶段2：模型训练

CCNN 干扰识别模型的定义、训练和推理。

## 目录结构

```
02_模型训练/
├── models/                     # 训练好的模型权重
│   ├── best/                  # 旧6类最佳 (99.13%)
│   ├── wide_jsr/              # 7类宽JSR (best: epoch_26 99.91%) ★最终使用
│   ├── archive/               # ONNX / .om 部署格式
│   └── legacy/                # 早期实验模型
│
├── scripts/                   # 代码
│   ├── training/
│   │   └── CCNN.py            # CCNN 模型定义 + 训练入口
│   ├── inference/
│   │   └── Infer.py           # 华为昇腾 NPU 推理引擎
│   └── utils/                 # 分析/对比/验证工具
│
├── results/                   # 训练日志与结果
│   ├── logs/                  # 推理测试日志
│   └── plots/                 # 训练曲线图
│
├── docs/                      # 实验报告
└── CCNN_Architecture_Detailed.md  # 架构详细说明
```

## CCNN 网络结构

```
输入 (2, 5000) IQ信号
  → CNN (Conv1d×3, 128ch) 局部时域特征
  → ResNet + SE注意力 (32ch) 通道加权
  → TCC 多头因果注意力 (4头) 全局时序依赖
  → 分类器 (7类: LFM/MTJ/NAM/NFM/STJ/SIN/Clean)
```

- 参数量：96,921
- 最佳准确率：99.91% (epoch 26)
- 误报率：0% (正常信号全部正确识别)

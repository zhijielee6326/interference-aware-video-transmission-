# 阶段1：数据生成

生成用于训练 CCNN 干扰识别模型的信号数据。

## 目录结构

```
01_数据生成/
├── MATLAB信号生成/              # MATLAB 版信号生成（旧版，保留对比）
│   ├── TX_LFM.m               # 线性扫频干扰
│   ├── TX_MTJ.m               # 多音干扰
│   ├── TX_NAM.m               # 窄带调幅干扰
│   ├── TX_NFM.m               # 窄带调频干扰
│   ├── TX_STJ.m               # 单音干扰
│   ├── TX_SIN.m               # 正弦波干扰
│   ├── PSK_MOD.m              # PSK 调制
│   ├── study.m                # 批量生成主脚本
│   ├── data_process.py        # .mat → .npz 格式转换
│   ├── generate_inference_testset.py  # 推理测试集生成
│   └── data/                  # 生成的 .mat 文件 (3600个)
│
├── generate_training_data.py  # Python 版批量生成宽JSR训练数据
├── usrp_model_test_v4.py      # Python 信号生成器（替代 MATLAB）
│
└── datasets/                  # 生成后的数据集
    ├── samples/               # 各类信号演示样本 (6 npz)
    ├── train/                 # 旧6类训练集
    ├── train_wide_jsr/        # 7类宽JSR训练集 (最佳)
    ├── test/                  # 测试集
    └── raw/                   # 原始捕获数据
```

## 两套生成方案对比

| 对比项 | MATLAB | Python |
|--------|--------|--------|
| 代码量 | 1457行 (16个.m) | 465行 (1个.py) |
| 干扰类型 | 6类 | 7类（含无干扰） |
| JSR范围 | 固定 6/8/10 dB | 可配置 -2~16 dB |
| 存储 | 3600个.mat (127MB) | 按需在线生成 |

Python 方案是最终使用的版本，MATLAB 保留用于对比验证。

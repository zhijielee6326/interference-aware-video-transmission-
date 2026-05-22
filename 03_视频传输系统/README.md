# 阶段3：视频传输系统集成

完整的干扰感知自适应视频传输系统，集成 CCNN 检测器实现实时干扰识别与传输参数动态调整。

## 核心文件

| 文件 | 功能 |
|------|------|
| `video_transmission_system.py` | **主系统** (1500+ 行)：视频采集→JPEG编码→FEC→QPSK调制→传输→解调→CCNN检测→自适应控制 |
| `cnn_interference_detector.py` | CCNN 检测器封装：加载模型、功率归一化、实时推理 |
| `run_system_with_viz.py` | 实时可视化启动器：4面板（频谱/概率/波形/JSR曲线） |
| `complete_system.py` | USRP B210 稳定收发系统（双线程解耦） |

## 启动脚本

| 脚本 | 用途 |
|------|------|
| `run_demo.sh` | 一键启动（支持 sim 仿真 / full 实机模式） |
| `run_loopback_test.sh` | USRP 自发自收快速测试 |
| `run_antenna_test.sh` | USRP 天线对天线测试 |

## USRP 硬件 Demo

| 文件 | 功能 |
|------|------|
| `usrp_chain_demo.py` | USRP 入门全链路 Demo（单音信号收发） |
| `usrp_gnuradio_antenna_demo.py` | GNU Radio + PyQt5 天线频谱演示 |
| `usrp_model_test_v4_fix.py` | USRP 模型测试 v4 修复版 |

## 信号生成（集成到系统内）

| 文件 | 功能 |
|------|------|
| `generate_training_data.py` | Python 版训练数据批量生成（从阶段1复制） |
| `matlab_signal_analysis.png` | MATLAB 信号分析参考图 |
| `training_data_inspect.png` | 训练数据检查参考图 |

## 自适应传输策略

| 干扰程度 | 触发条件 | JPEG质量 | FEC冗余 | 帧率 |
|----------|----------|----------|---------|------|
| 无干扰 | Clean | 70 | 1x | 25 fps |
| 轻度 | NAM/NFM/SIN | 50 | 2x | 15 fps |
| 中度 | LFM/MTJ | 30 | 2x | 8 fps |
| 严重 | STJ | 20 | 3x | 5 fps |

> 滞后计数器：干扰连续5次确认切换，无干扰连续8次恢复，防止策略抖动。

## 说明文档

各模块详细说明见 `README_*.md` 文件。

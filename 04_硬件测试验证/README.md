# 阶段4：硬件测试与验证

USRP B210 硬件在环测试及系统验证实验。

## 目录内容

### USRP 实机测试

| 文件 | 功能 |
|------|------|
| `usrp_model_test.py` | 实机 JSR 扫描测试脚本 |
| `plot_usrp_jsr.py` | JSR 曲线绘图脚本 |
| `usrp_jsr_results.json` | 实测 JSR 数据 |
| `usrp_jsr_accuracy.pdf` / `.png` | JSR 精度报告 |
| `usrp_test_report.png` | USRP 测试报告截图 |
| `jsr_accuracy.png` | JSR 准确率曲线 |

### JSR 扫描测试（仿真）

| 文件 | 功能 |
|------|------|
| `jsr_sweep_test.py` | JSR 精度扫描测试脚本 |
| `jsr_sweep_results.json` | 扫描结果数据 |
| `jsr_sweep_results.png` | 扫描结果图 |
| `jsr_results.json` | JSR 测试结果 |
| `jsr_sweep_correct.json` | 修正后的 JSR 扫描结果 |
| `experiment_jammer_results.json` | 干扰注入实验数据 |

### 离线测试

| 文件 | 功能 |
|------|------|
| `test_ccnn.py` | CCNN 7类离线准确率测试 |
| `test_adaptive.py` | 自适应传输策略验证 |
| `test_detector_simple.py` | 检测器快速测试 |
| `test_signal_length.py` | 信号长度诊断 |
| `adaptive_test_log.txt` | 自适应测试日志 |
| `test_report.txt` | 测试报告 |

### 论文实验图表

| 文件 | 出处 |
|------|------|
| `figure_2_1_qpsk_constellation.png` | 论文图2-1：QPSK星座图 |
| `figure_4_2_interference_waveforms.png` | 论文图4-2：干扰波形 |
| `figure_4_4_training_curve.png` | 论文图4-4：训练曲线 |
| `figure_4_5_confusion_matrix.png` | 论文图4-5：混淆矩阵 |
| `fig_adaptive_comparison.png` | 自适应策略对比图 |
| `fig_adaptive_transmission.png` | 自适应传输效果图 |
| `fig_psd_comparison.png` | PSD 对比图 |
| `gen_thesis_figures.py` | 论文图表生成脚本 |

## 关键实验结果

- **仿真 JSR 扫描**: -2dB 时仍达 99.4%，0dB 以上达 99.9%
- **USRP 硬件测试**: 干扰注入后检测有效，自适应策略正确切换，清除后快速恢复

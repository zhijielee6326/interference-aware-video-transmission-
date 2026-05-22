#!/usr/bin/env python3
"""
CCNN七类干扰识别离线测试脚本。

对6类干扰和无干扰各生成10个样本，输出分类准确率和混淆矩阵，
并将结果保存到 gnuradio/test_report.txt。
"""
import argparse
import os
import sys
import types
from datetime import datetime

import numpy as np

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.abspath(os.path.join(SCRIPT_DIR, '..'))
MODEL_PATH = os.path.join(
    PROJECT_ROOT, 'CCNN', '2_models', 'wide_jsr',
    'ccnn_epoch_26_acc_0.9991.pth')
CCNN_PATH = os.path.join(
    PROJECT_ROOT, 'CCNN', '3_scripts', 'training', 'CCNN.py')
REPORT_PATH = os.path.join(SCRIPT_DIR, 'test_report.txt')

# usrp_model_test_v4.py 顶层会导入uhd；离线生成信号不需要真实USRP。
if 'uhd' not in sys.modules:
    sys.modules['uhd'] = types.SimpleNamespace()

sys.path.insert(0, SCRIPT_DIR)
from cnn_interference_detector import InterferenceDetector
from usrp_model_test_v4 import FS, RS, T, gen_psk, make_combined

CLASSES = ['lfm', 'mtj', 'nam', 'nfm', 'stj', 'sin', 'clean']
LABELS = [
    '扫频干扰(LFM)',
    '多音干扰(MTJ)',
    '窄带AM(NAM)',
    '窄带FM(NFM)',
    '单音干扰(STJ)',
    '正弦波(SIN)',
    '无干扰',
]


def make_sample(class_key: str, jsr_db: float) -> np.ndarray:
    if class_key == 'clean':
        return gen_psk('QPSK', RS, 0, FS, T).astype(np.complex64)
    return make_combined(class_key, jsr_db=jsr_db, modulation='QPSK')


def format_matrix(matrix: np.ndarray) -> str:
    rows = ['真实\\预测 ' + ' '.join(f'{name:^8}' for name in CLASSES)]
    for key, row in zip(CLASSES, matrix):
        rows.append(f'{key:>8} ' + ' '.join(f'{v:^8d}' for v in row))
    return '\n'.join(rows)


def main():
    parser = argparse.ArgumentParser(description='CCNN七类干扰离线测试')
    parser.add_argument('--samples', type=int, default=10, help='每类测试样本数')
    parser.add_argument('--jsr', type=float, default=10.0, help='干信比(dB)')
    parser.add_argument('--model', default=MODEL_PATH, help='模型权重路径')
    parser.add_argument('--report', default=REPORT_PATH, help='报告输出路径')
    args = parser.parse_args()

    np.random.seed(2026)
    detector = InterferenceDetector(
        model_path=args.model,
        ccnn_py=CCNN_PATH,
        num_classes=7,
    )

    matrix = np.zeros((len(CLASSES), len(CLASSES)), dtype=int)
    lines = [
        'CCNN七类干扰识别离线测试报告',
        f'生成时间：{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}',
        f'模型文件：{args.model}',
        f'样本设置：每类{args.samples}个，干信比JSR={args.jsr:.1f} dB',
        '',
        '逐类结果：',
    ]

    for true_idx, class_key in enumerate(CLASSES):
        correct = 0
        for _ in range(args.samples):
            sample = make_sample(class_key, args.jsr)
            pred_label, conf = detector.detect(sample)
            pred_idx = LABELS.index(pred_label) if pred_label in LABELS else -1
            if pred_idx >= 0:
                matrix[true_idx, pred_idx] += 1
            correct += int(pred_idx == true_idx)
        acc = correct / args.samples
        lines.append(
            f'- {class_key.upper():>5}（{LABELS[true_idx]}）：'
            f'{correct}/{args.samples}，准确率 {acc:.2%}')

    total_correct = int(np.trace(matrix))
    total = int(matrix.sum())
    overall = total_correct / total if total else 0.0
    lines.extend([
        '',
        f'总体准确率：{total_correct}/{total} = {overall:.2%}',
        '',
        '混淆矩阵：',
        format_matrix(matrix),
    ])

    report = '\n'.join(lines) + '\n'
    with open(args.report, 'w', encoding='utf-8') as f:
        f.write(report)
    print(report)
    print(f'报告已保存：{args.report}')


if __name__ == '__main__':
    main()

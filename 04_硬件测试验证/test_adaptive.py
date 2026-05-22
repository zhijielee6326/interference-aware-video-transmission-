#!/usr/bin/env python3
"""
自适应传输策略验证脚本。

验证8帧确认干扰、3帧确认等级、12帧恢复的滞后逻辑，
并输出JPEG质量、FEC冗余和帧率切换时序。
"""
import os
import sys
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

from video_transmission_system import INTERFERENCE_PROFILES, TransmissionController

LOG_PATH = os.path.join(SCRIPT_DIR, 'adaptive_test_log.txt')


def run_sequence(controller, name, detections):
    rows = []
    last_state = controller.current_severity
    for frame_idx, (itype, conf) in enumerate(detections, start=1):
        controller.update(itype, conf)
        params = controller.get_smoothed_params()
        changed = controller.current_severity != last_state
        rows.append({
            'frame': frame_idx,
            'input': itype,
            'conf': conf,
            'severity': controller.current_severity,
            'quality': params['quality'],
            'fec': params['fec_redundancy'],
            'fps': params['fps'],
            'changed': changed,
        })
        last_state = controller.current_severity
    return name, rows


def assert_first_change(rows, expected_frame, expected_severity):
    changes = [r for r in rows if r['changed']]
    if not changes:
        raise AssertionError(f'未发生切换，期望第{expected_frame}帧切到{expected_severity}')
    first = changes[0]
    if first['frame'] != expected_frame or first['severity'] != expected_severity:
        raise AssertionError(
            f'切换不符合预期：实际第{first["frame"]}帧->{first["severity"]}，'
            f'期望第{expected_frame}帧->{expected_severity}')


def format_rows(rows):
    lines = ['帧号 | 输入干扰 | 置信度 | 策略等级 | JPEG质量 | FEC | FPS | 事件']
    lines.append('--- | --- | --- | --- | --- | --- | --- | ---')
    for row in rows:
        event = '切换' if row['changed'] else ''
        lines.append(
            f'{row["frame"]} | {row["input"]} | {row["conf"]:.2f} | '
            f'{row["severity"]} | {row["quality"]} | {row["fec"]} | '
            f'{row["fps"]} | {event}')
    return '\n'.join(lines)


def main():
    report = [
        '自适应传输策略验证报告',
        f'生成时间：{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}',
        '',
        '策略参数：',
    ]
    for key, profile in INTERFERENCE_PROFILES.items():
        report.append(
            f'- {key}: JPEG质量={profile["quality"]}, '
            f'FEC={profile["fec_redundancy"]}, FPS={profile["fps"]}')

    scenarios = []

    severe_controller = TransmissionController()
    severe_seq = [('单音干扰(STJ)', 0.95)] * 12
    name, rows = run_sequence(severe_controller, '连续严重干扰确认', severe_seq)
    assert_first_change(rows, 10, 'severe')
    scenarios.append((name, rows, '通过：第8帧确认干扰，第10帧完成3帧等级确认并切到severe。'))

    controller = TransmissionController()
    mild_seq = [('窄带AM(NAM)', 0.90)] * 12
    name, rows = run_sequence(controller, '首次轻度干扰确认', mild_seq)
    assert_first_change(rows, 10, 'mild')
    scenarios.append((name, rows, '通过：首次出现干扰时仍需8帧干扰确认和3帧等级确认。'))

    mild_seq = [('窄带AM(NAM)', 0.90)] * 3
    name, rows = run_sequence(severe_controller, '干扰等级切换确认', mild_seq)
    if rows[-1]['severity'] != 'mild':
        raise AssertionError('已确认干扰状态下，连续3帧mild后未切换到mild')
    scenarios.append((name, rows, '通过：干扰已确认后，连续同一mild等级3帧即可切换，参数跟随mild档。'))

    recover_seq = [('无干扰', 0.99)] * 12
    name, rows = run_sequence(severe_controller, '无干扰恢复确认', recover_seq)
    assert_first_change(rows, 12, 'none')
    scenarios.append((name, rows, '通过：连续12帧无干扰后恢复none档。'))

    for name, rows, conclusion in scenarios:
        report.extend(['', f'## {name}', format_rows(rows), conclusion])

    final = '\n'.join(report) + '\n'
    with open(LOG_PATH, 'w', encoding='utf-8') as f:
        f.write(final)
    print(final)
    print(f'日志已保存：{LOG_PATH}')


if __name__ == '__main__':
    main()

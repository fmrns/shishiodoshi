#!/usr/bin/env python3
# -*- coding: utf-8 -*-
#
# Released under MIT License
#
# Copyright (c) 2025 Fumiyuki Shimizu
# Copyright (c) 2025 Abacus Technologies, Inc.
#
# Permission is hereby granted, free of charge, to any person
# obtaining a copy of this software and associated documentation files
# (the "Software"), to deal in the Software without restriction,
# including without limitation the rights to use, copy, modify, merge,
# publish, distribute, sublicense, and/or sell copies of the Software,
# and to permit persons to whom the Software is furnished to do so,
# subject to the following conditions:
#
# The above copyright notice and this permission notice shall be
# included in all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
# MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
# NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS
# BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
# ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
# CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

import sys
import os

sys.path.insert(0, os.path.dirname(__file__))

import subprocess
from datetime import datetime as datetime_sucks
from pendulum import (
    DateTime,
    datetime,
    now,
    duration as penduration,
    parse as penparse,
    SATURDAY,
    SUNDAY,
)
from openpyxl import load_workbook

from timerange import TimeRange, TimeRangeSet
from task import Task, TaskSet
from member import Member, MemberSet

subprocess.run("c:/Windows/System32/mode.com con cp select=65001", shell=True)
# print(f"sys.stdout.encoding: {sys.stdout.encoding}")

THE_FIRST_DATE = datetime(2025, 9, 16)
THE_LAST_DATE = datetime(2025, 10, 1)


def make_breaks():
    break_start1 = THE_FIRST_DATE.add(days=-1).at(17)
    break_end1 = THE_FIRST_DATE.at(9)
    break_start2 = THE_FIRST_DATE.at(12)
    break_end2 = THE_FIRST_DATE.at(13)

    breaks = TimeRangeSet()
    # 祝日
    breaks.add(TimeRange(datetime(2025, 9, 23, 0), datetime(2025, 9, 24, 0)))
    # 平時、土日
    while break_start1 <= THE_LAST_DATE.add(days=1).at(0):
        breaks.add(TimeRange(break_start1, break_end1))
        breaks.add(TimeRange(break_start2, break_end2))
        if break_end1.day_of_week in [SATURDAY, SUNDAY]:
            breaks.add(TimeRange(break_end1.at(0), break_end1.add(days=1).at(0)))
        break_start1 = break_start1.add(days=1)
        break_end1 = break_end1.add(days=1)
        break_start2 = break_start2.add(days=1)
        break_end2 = break_end2.add(days=1)
    return breaks


def load_members(ws) -> MemberSet:
    labels = [cell.value for cell in ws[1]]
    label_to_col = {label: i for i, label in enumerate(labels, start=1) if label}

    errors = []
    baseline = ws.cell(row=2, column=label_to_col["ベースライン"]).value
    team = ws.cell(row=2, column=label_to_col["チーム名"]).value

    members = MemberSet()
    for i in range(1, 10):
        col = label_to_col.get(f"担当者{i}")
        if col is not None:
            name = ws.cell(row=2, column=col).value
            if name is not None:
                role = ws.cell(row=3, column=col).value or "プログラマ"
                members.add(Member(name=name, role=role))
    if baseline is None:
        errors.append("ベースラインの定義が見つかりません。しくしく...")
    if team is None:
        errors.append("チーム名の定義が見つかりません。しくしく...")
    if 1 > len(members):
        errors.append("担当者の定義が見当たりません。しくしく...")
    if len(errors) > 0:
        raise ValueError(errors)

    return baseline, team, members


def to_datetime(val):
    if val is None:
        return val
    if isinstance(val, DateTime):
        return val
    if isinstance(val, str):
        try:
            return datetime.from_format(val, "YYYY-MM-DD HH:mm:ss")
        except Exception:
            return None
    if isinstance(val, datetime_sucks):
        return datetime(
            val.year,
            val.month,
            val.day,
            val.hour,
            val.minute,
            val.second,
        )
    return val


def load_tasks(ws, members: MemberSet, breaks: TimeRangeSet, now: DateTime) -> TaskSet:
    labels = None
    for row in ws.iter_rows(min_row=5, max_row=ws.max_row, values_only=True):
        if "タスク" in row:
            labels = list(row)
            break
    if labels is None:
        raise ValueError("'タスク'というセルが見つかりません。しくしく...")
    label_to_col = {label: i for i, label in enumerate(labels) if label}
    # print(f"{label_to_col}")

    if "進捗\n(%)" in label_to_col and "進捗(%)" not in label_to_col:
        label_to_col["進捗(%)"] = label_to_col["進捗\n(%)"]
    for label in [
        "タスク",
        "進捗(%)",
        "予定開始日時",
        "予定完了日時",
        "実績開始日時",
        "実績完了日時",
        "担当者1",
    ]:
        if label not in label_to_col:
            raise ValueError(f"{label}列が見つかりません。しくしく...")
        # else:
        #    print(f"{label}: 列{label_to_col[label]}")

    taskset = TaskSet()
    is_warned = False
    for i, row in enumerate(
        ws.iter_rows(min_row=6, max_row=ws.max_row, values_only=True), start=6
    ):
        task_name = row[label_to_col["タスク"]]
        if task_name is None:
            continue
        task_name += f"-line{i}"
        plan_start = to_datetime(row[label_to_col["予定開始日時"]])
        if plan_start is None:
            continue
        plan_end = to_datetime(row[label_to_col["予定完了日時"]])
        if plan_end is None:
            print(f"{task_name}: 予定完了日時がありません。")
        try:
            progress = int(row[label_to_col["進捗(%)"]])
        except (TypeError, ValueError):
            progress = -1
        actual_start = to_datetime(row[label_to_col["実績開始日時"]])
        actual_end = to_datetime(row[label_to_col["実績完了日時"]])
        task_members = MemberSet()
        for j in range(1, 10):
            col = label_to_col.get(f"担当者{j}")
            if col is not None:
                name = row[col]
                if name is not None:
                    m = members.find_by_name(name)
                    if m is not None:
                        task_members.add(m)
                    else:
                        print(
                            f"{task_name}: 担当者{j}の{name}さんは定義されていません。"
                        )
        if len(task_members) < 1:
            is_warned = True
            print(f"{task_name}: 恐ろしいことに誰も担当していません。")
        task = Task(
            task_name,
            progress,
            plan_start,
            plan_end,
            actual_start,
            actual_end,
            now,
            breaks,
        )
        if task.was_warned:
            is_warned = True
        taskset.add(task)
        for m in task_members:
            m.add_task(task)
            if m.was_warned:
                is_warned = True
    if is_warned:
        print("\n確認したらenterを押してください。")
        input()
    return taskset


def print_progress_details(
    planned_total,
    planned_done,
    actual_done,
    expected_actual_seconds,
    note="",
):
    planned_progress = planned_done.in_seconds() / planned_total.in_seconds()
    actual_progress = actual_done.in_seconds() / expected_actual_seconds
    actual_per_planned = (
        actual_progress / planned_progress if planned_progress > 0 else -1
    )
    print(
        f"予定進捗率{note}: {100 * planned_progress:.2f}% "
        f"({planned_done.in_hours()}hr/{planned_total.in_hours()}hr)"
    )
    print(
        f"実績進捗率{note}: {100 * actual_progress:.2f}% "
        f"({actual_done.in_hours()}hr/{expected_actual_seconds / 3600.0:.0f}hr)"
    )
    print(
        f"実績/予定 {note}: "
        + (f"{100 * actual_per_planned:.2f}%" if actual_per_planned >= 0 else "N/A")
        + f" ({100 * actual_progress:.2f}%/{100 * planned_progress:.2f}%)"
    )
    print(
        "コメント  : "
        + (
            "がんばります。"
            if actual_per_planned < 50
            else "だんだん調子が出てきました。"
            if actual_per_planned < 90
            else "順調です。"
            if actual_per_planned < 120
            else "順調すぎて怖いです。"
        )
    )


def main():
    breaks = make_breaks()

    if len(sys.argv) > 1:
        wb = load_workbook(sys.argv[1])
    else:
        wb = load_workbook("進捗管理表.xlsx")

    if len(sys.argv) > 2:
        nowt = penparse(sys.argv[2])
    else:
        nowt = now()

    ws = wb.active
    baseline, team, members = load_members(ws)
    print(f"基準: {baseline}")
    print(f"チーム名: {team}")
    for m in members:
        print("-" * 50)
        print(f"{m.name}さん")
        print(f"役割: {m.role}")
    print("-" * 50)

    load_tasks(ws, members, breaks, nowt)
    # for task in tasks:
    #     print("-" * 50)
    #     print(f"タスク: {task.name}")
    #     print(f"進捗: {task.progress}%")
    #     print(f"予定: {task.plan_start} - {task.plan_end}")
    #     print(f"実績: {task.actual_start} - {task.actual_end}")

    whole_period = THE_LAST_DATE.add(days=1).at(0).diff(THE_FIRST_DATE.at(0))
    passed_period = max(
        penduration(), nowt.add(days=1).at(0).diff(THE_FIRST_DATE.at(0))
    )
    team_planned_total = penduration()
    team_planned_done = penduration()
    team_actual_done = penduration()
    team_expected_actual_seconds = 0
    for m in members:
        planned_total, planned_done, actual_done, expected_actual_seconds = (
            m.tasks.total_durations()
        )
        team_planned_total += planned_total
        team_planned_done += planned_done
        team_actual_done += actual_done
        team_expected_actual_seconds += expected_actual_seconds

    is_team_shown = False
    for m in members:
        print("=" * 80)
        print(f"{m.name}さん")
        print("-" * 50)
        print(f"チーム名  : {team}")
        print(f"担当タスク: {', '.join(m.tasks.names_responsible(nowt))}")
        print(f"役割      : {m.role}")
        print(f"ベースライン: {baseline}")
        print(
            f"行程      : {100 * passed_period.in_days() / whole_period.in_days():.2f}% "
            f"({passed_period.in_days()}日/{whole_period.in_days()}日)"
        )

        planned_total, planned_done, actual_done, expected_actual_seconds = (
            m.tasks.total_durations()
        )
        print_progress_details(
            planned_total,
            planned_done,
            actual_done,
            expected_actual_seconds,
        )
        if m.role == "リーダー":
            is_team_shown = True
            print("-" * 50)
            print_progress_details(
                team_planned_total,
                team_planned_done,
                team_actual_done,
                team_expected_actual_seconds,
                "(チーム)",
            )

    if not is_team_shown:
        print("=" * 80)
        print_progress_details(
            team_planned_total,
            team_planned_done,
            team_actual_done,
            team_expected_actual_seconds,
            "(チーム)",
        )


if __name__ == "__main__":
    main()

# end of file

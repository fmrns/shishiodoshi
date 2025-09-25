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

from pendulum import DateTime, duration as penduration
from timerange.timerange import TimeRange
from timerange.timerangeset import TimeRangeSet


class Task:
    def __init__(
        self,
        name: str,
        line: int,
        progress: int,
        plan_start: DateTime,
        plan_end: DateTime,
        actual_start: DateTime | None,
        actual_end: DateTime | None,
        now: DateTime,
        breaks: TimeRangeSet,
    ):
        self.name = f"{name}-line{line}"
        self.plan_start = plan_start
        self.plan_end = plan_end
        self.actual_start = actual_start
        self.actual_end = actual_end
        self.progress = progress
        self.now = now
        self.was_warned = False
        self._validate()
        self.plan = TimeRangeSet([TimeRange(self.plan_start, self.plan_end)]) - breaks
        if self.actual_end is not None:
            self.actual = (
                TimeRangeSet([TimeRange(self.actual_start, self.actual_end)]) - breaks
            )
        elif self.actual_start is not None and self.actual_start < self.now:
            self.actual = (
                TimeRangeSet([TimeRange(self.actual_start, self.now)]) - breaks
            )
        else:
            self.actual = None
        self.planned_total_duration = self.plan.total_duration()
        self.planned_done_duration = (
            TimeRangeSet([TimeRange(self.plan_start, min(self.now, self.plan_end))])
            - breaks
            if self.plan_start < self.now
            else None
        )
        if self.planned_done_duration is None:
            self.planned_done_duration = penduration()
        self.actual_done_duration = (
            self.actual.total_duration() if self.actual is not None else penduration()
        )
        # 実績がある場合は工数を進捗率から予測。
        self.expected_actual_total_seconds = (
            self.actual_done_duration.in_seconds() * 100.0 / self.progress
            if self.progress > 0
            else self.planned_total_duration.in_seconds()
        )

    def __eq__(self, val: "Task"):
        return self.name == val.name

    def is_overlap(self, other: "Task") -> bool:
        return TimeRange(self.plan_start, self.plan_end).is_overlap(
            TimeRange(other.plan_start, other.plan_end)
        )

    def _validate(self):
        if self.plan_start == self.plan_end:
            self.plan_end = self.plan_end.add(seconds=1)
            self.was_warned = True
            print(
                f"{self.name}: 予定開始日時が予定終了日時と同じです。時間指定を間違えないように注意して修正してください。"
            )
        elif self.plan_start > self.plan_end:
            t = self.plan_start
            self.plan_start = self.plan_end
            self.plan_end = t
            self.was_warned = True
            print(
                f"{self.name}: 予定開始日時が予定終了日時より後です。時間指定を間違えないように注意して修正してください。"
            )
        if self.actual_start is None:
            if self.actual_end is not None:
                self.actual_start = self.actual_end
                self.actual_end = None
                self.was_warned = True
                print(
                    f"{self.name}: 実績開始日時がないのに実績終了日時があります。修正してください。"
                )
        elif self.actual_end is not None:
            if self.actual_start == self.actual_end:
                self.actual_end = self.actual_end.add(seconds=1)
                self.was_warned = True
                print(
                    f"{self.name}: 実績開始日時が実績終了日時と同じです。時間指定を間違えないように注意して修正してください。"
                )
            elif self.actual_start > self.actual_end:
                t = self.actual_start
                self.actual_start = self.actual_end
                self.actual_end = t
                self.was_warned = True
                print(
                    f"{self.name}: 実績開始日時が実績終了日時より後です。時間指定を間違えないように注意して修正してください。"
                )
        if (
            self.actual_start is not None
            and self.actual_end is None
            and self.actual_start > self.now
            and self.progress > 0
        ):
            self.actual_start = None
            self.progress = 0
            self.was_warned = True
            print(f"{self.name}: 実績開始日時が未来です。")
        if self.actual_end is not None and self.actual_end > self.now:
            self.was_warned = True
            print(f"{self.name}: 実績完了日時が未来です。")
        if self.actual_start is None and self.progress > 0:
            self.actual_start = self.plan_start
            self.was_warned = True
            print(f"{self.name}: 進捗があります。実績開始日時をセットしてください。")
        if self.progress > 100:
            print(f"{self.name}: 進捗しすぎ({self.progress}%)です。")
            self.progress = 100
            self.was_warned = True
        if self.actual_end is not None and self.progress != 100:
            self.progress = 100
            self.was_warned = True
            print(
                f"{self.name}: 実績完了日時があります。進捗を100%にセットしてください。"
            )
        if self.actual_start is not None and self.progress < 0:
            self.progress = 0


# end of file

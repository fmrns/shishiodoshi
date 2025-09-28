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
        progress: int,
        plan_start: DateTime,
        plan_end: DateTime,
        actual_start: DateTime | None,
        actual_end: DateTime | None,
        now: DateTime,
        breaks: TimeRangeSet,
        on_off_map: dict[DateTime, bool],
    ):
        self.name = name
        self.plan_start = plan_start
        self.plan_end = plan_end
        self.actual_start = actual_start
        self.actual_end = actual_end
        self.progress = progress
        self.now = now
        self.was_warned = False
        self._validate(on_off_map)

        trs = TimeRangeSet([TimeRange(self.plan_start, self.plan_end)]) - breaks
        dur = trs.total_duration() if trs else penduration()
        self.planned_total_seconds = dur.in_seconds()

        trs = (
            TimeRangeSet([TimeRange(self.plan_start, min(self.now, self.plan_end))])
            - breaks
            if self.plan_start < self.now
            else None
        )
        dur = trs.total_duration() if trs else penduration()
        self.planned_done_seconds = dur.in_seconds()

        if self.progress == 100:
            trs = TimeRangeSet([TimeRange(self.actual_start, self.actual_end)]) - breaks
            dur = trs.total_duration() if trs else penduration()
            self.actual_total_seconds = dur.in_seconds()
            self.actual_done_seconds = self.actual_total_seconds
        elif self.progress >= 30:
            # 工数を進捗率から予測。
            trs = TimeRangeSet([TimeRange(self.actual_start, self.now)]) - breaks
            dur = trs.total_duration() if trs else penduration()
            self.actual_done_seconds = dur.in_seconds()
            self.actual_total_seconds = self.actual_done_seconds * 100 / self.progress
        else:
            # 誤差が大きくなるので工数は計画を利用
            self.actual_total_seconds = self.planned_total_seconds
            self.actual_done_seconds = self.actual_total_seconds * self.progress / 100

    def __eq__(self, val: "Task"):
        return self.name == val.name

    def is_overlap(self, other: "Task") -> bool:
        return TimeRange(self.plan_start, self.plan_end).is_overlap(
            TimeRange(other.plan_start, other.plan_end)
        )

    def period_start(self) -> DateTime:
        return (
            self.actual_start
            if self.actual_start and self.actual_start < self.plan_start
            else self.plan_start
        )

    def period_end(self) -> DateTime:
        return (
            self.actual_end
            if self.actual_end and self.actual_end > self.plan_end
            else self.plan_end
        )

    # vvvvvvvvvvvvvvvvvvvvvvvvvvvvvv フィルター vvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
    def is_responsible(self, base_date: DateTime) -> bool:
        base_start = base_date.at(0)
        base_end = base_date.add(days=1)
        if self.actual_start is None:
            return self.plan_start < base_end
        return not self.actual_end or base_start <= self.actual_end

    def is_unstarted(self, nowt: DateTime) -> bool:
        return not self.actual_start and self.plan_start < nowt

    def is_unfinished(self, nowt: DateTime) -> bool:
        return self.actual_start and self.plan_end <= nowt

    def is_overrun(self, nowt: DateTime) -> bool:
        return (
            nowt - self.actual_start > self.plan_end - self.plan_start
            if not self.actual_end and self.actual_start
            else False
        )

    # ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ フィルター ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

    def _validate(self, on_off_map: dict[DateTime, bool]):
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
            if self.actual_end:
                self.actual_start = self.actual_end
                self.actual_end = None
                self.was_warned = True
                print(
                    f"{self.name}: 実績開始日時がないのに実績終了日時があります。修正してください。"
                )
        elif self.actual_end:
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

        if self.progress:
            if self.progress > 100:
                print(f"{self.name}: 進捗よすぎ({self.progress}%)です。")
                self.progress = 100
                self.was_warned = True
            elif self.progress < 0:
                print(f"{self.name}: 進捗とんでもなく悪すぎ({self.progress}%)です。")
                self.progress = 0
                self.was_warned = True
            if not self.actual_start and self.progress > 0:
                self.actual_start = self.plan_start
                self.was_warned = True
                print(
                    f"{self.name}: 進捗があります。実績開始日時をセットしてください。"
                )
        else:
            self.progress = 0
            if (
                not self.actual_end
                and self.actual_start
                and self.actual_start < self.now
            ):
                self.was_warned = True
                print(
                    f"{self.name}: 実績開始日時があります。進捗をセットしてください。"
                )
        if self.actual_end and self.progress != 100:
            self.progress = 100
            self.was_warned = True
            print(
                f"{self.name}: 実績完了日時があります。進捗を100%にセットしてください。"
            )

        if (
            self.actual_start
            and not self.actual_end
            and self.actual_start > self.now
            and self.progress > 0
        ):
            self.actual_start = None
            self.progress = 0
            self.was_warned = True
            print(f"{self.name}: 実績開始日時が未来です。")
        if self.actual_end and self.actual_end > self.now:
            self.was_warned = True
            print(f"{self.name}: 実績完了日時が未来です。")

        if self.actual_start and self.progress < 0:
            self.progress = 0

        dt = self.plan_start.at(0)
        while dt <= self.plan_end:
            if dt not in on_off_map:
                raise ValueError(
                    f"{dt.format('YYYY-MM-DD HH:mm')}がカレンダー(O列以降)にありません。"
                )
            dt = dt.add(days=1)


# end of file

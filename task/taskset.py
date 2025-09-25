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

from pendulum import Duration, duration as penduration
from task.task import Task


class TaskSet:
    def __init__(self, tasks: list[Task] = None):
        self.tasks = tasks or []

    def add(self, task: Task):
        if task not in self.tasks:
            self.tasks.append(task)

    def __iter__(self):
        return iter(self.tasks)

    def __len__(self):
        return len(self.tasks)

    def __getitem__(self, index):
        return self.tasks[index]

    def names(self) -> list[str]:
        return [task.name for task in self.tasks]

    def total_planned_duration(self) -> Duration:
        return sum((t.planned_total_duration for t in self.tasks), penduration())

    def total_planned_done_duration(self) -> Duration:
        return sum((t.planned_done_duration for t in self.tasks), penduration())

    def total_actual_done_duration(self) -> Duration:
        return sum((t.actual_done_duration for t in self.tasks), penduration())

    def total_expected_actual_seconds(self) -> float:
        return sum(t.expected_actual_total_seconds for t in self.tasks)

    def total_durations(self) -> tuple[Duration, Duration, Duration, float]:
        planned_total = penduration()
        planned_done = penduration()
        actual_done = penduration()
        expected_actual_seconds = 0.0
        for t in self.tasks:
            planned_total += t.planned_total_duration
            planned_done += t.planned_done_duration
            actual_done += t.actual_done_duration
            expected_actual_seconds += t.expected_actual_total_seconds
        return planned_total, planned_done, actual_done, expected_actual_seconds


# end of file

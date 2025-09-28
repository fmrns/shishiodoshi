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

from typing import Callable
from pendulum import DateTime, Duration, duration as penduration
from util.text import wlen
from task.task import Task


class TaskSet:
  def __init__(self, tasks: list[Task] = None):
    self.tasks = tasks or []
    if self.tasks:
      self.period_start = min(t.period_start() for t in self.tasks)
      self.period_end = max(t.period_end() for t in self.tasks)
      self.max_len_of_names = max(wlen(t.name) for t in self.tasks)
    else:
      self.period_start = None
      self.period_end = None
      self.max_len_of_names = None

  def add(self, task: Task):
    if task not in self.tasks:
      self.tasks.append(task)
      ltn = wlen(task.name)
      if not self.max_len_of_names or self.max_len_of_names < ltn:
        self.max_len_of_names = ltn
      tps = task.period_start()
      tpe = task.period_end()
      if not self.period_start or self.period_start > tps:
        self.period_start = tps
      if not self.period_end or self.period_end < tpe:
        self.period_end = tpe

  def add_tasks(self, ts: "TaskSet"):
    for t in ts:
      self.add(t)

  def __iter__(self):
    return iter(self.tasks)

  def __len__(self):
    return len(self.tasks)

  def __getitem__(self, index):
    return self.tasks[index]

  def names(self) -> list[str]:
    return [task.name for task in self.tasks]

  def filter(self, pred: Callable[[Task], bool]) -> "TaskSet":
    return TaskSet([t for t in self.tasks if pred(t)])

  def calc_base(self, base_date: DateTime) -> tuple["DateTime", "DateTime"]:
    if not self.period_start or self.period_start <= base_date <= self.period_end:
      rc = base_date.at(0)
    elif self.period_end < base_date:
      rc = self.period_end.at(0)
    else:
      rc = self.period_start.at(0)
    return rc, rc.add(days=1)

  def total_durations(self) -> tuple[Duration, Duration, Duration, float]:
    planned_total_seconds = 0
    planned_done_seconds = 0
    actual_total_seconds = 0
    actual_done_seconds = 0
    for t in self.tasks:
      planned_total_seconds += t.planned_total_seconds
      planned_done_seconds += t.planned_done_seconds
      actual_total_seconds += t.actual_total_seconds
      actual_done_seconds += t.actual_done_seconds
    return (
      planned_total_seconds,
      planned_done_seconds,
      actual_total_seconds,
      actual_done_seconds,
    )


# end of file

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

from task.task import Task
from task.taskset import TaskSet


class Member:
    def __init__(self, name: str, role: str | None):
        self.name = name
        self.role = role
        self.tasks = TaskSet()
        self.was_warned = False

    def __eq__(self, val: "Member"):
        return self.name == val.name

    def add_task(self, task: "Task"):
        for ts in self.tasks:
            if ts.is_overlap(task):
                self.was_warned = True
                print(
                    f"{self.name}さんのタスク「{task.name}」はタスク「{ts.name}」と重なっています。タスクを分割するなどして修正してください。"
                )
        self.tasks.add(task)

    def __repr__(self):
        return f"Member(name={self.name!r}, role={self.role!r}, tasks={len(self.tasks)} tasks)"


# end of file

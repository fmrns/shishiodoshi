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

from pendulum import DateTime, Duration, timezone


class TimeRange:
  def __init__(self, start: DateTime, end: DateTime):
    if start >= end:
      raise ValueError("start must be before end")
    self.start = start
    self.end = end
    # TODO: failsafe for now. remove the following after a suitable grace period has passed.
    if start.tzinfo != timezone("Asia/Tokyo"):
      raise ValueError(f"tz of start is {start.tzinfo}")
    if end.tzinfo != timezone("Asia/Tokyo"):
      raise ValueError(f"tz of end is {end.tzinfo}")

  def is_overlap(self, other: "TimeRange") -> bool:
    return self.start < other.end and other.start < self.end

  def __sub__(self, other: "TimeRange") -> "TimeRangeSet":
    from timerange.timerangeset import TimeRangeSet

    rc = TimeRangeSet()
    if not self.is_overlap(other):
      rc.add(self)
    else:
      if self.start < other.start:
        rc.add(TimeRange(self.start, min(self.end, other.start)))
      if other.end < self.end:
        rc.add(TimeRange(max(self.start, other.end), self.end))
    return rc

  def duration(self) -> Duration:
    return self.end.diff(self.start)

  def __repr__(self):
    return (
      f"TimeRange({self.start.to_datetime_string()} - {self.end.to_datetime_string()})"
    )


# end of file

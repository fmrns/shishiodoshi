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

from timerange.timerange import TimeRange
from pendulum import Duration, duration as penduration


class TimeRangeSet:
    def __init__(self, ranges: list[TimeRange] = []):
        self.ranges = self._normalize(ranges or [])

    def add(self, r: TimeRange):
        self.ranges = self._normalize(self.ranges + [r])

    def __sub__(self, subtractors: "TimeRangeSet") -> "TimeRangeSet":
        rc = TimeRangeSet()
        for r in self.ranges:
            for s in subtractors:
                for rr in r - s:
                    rc.add(rr)
        return rc

    def total_duration(self) -> Duration:
        return sum((r.duration() for r in self.ranges), penduration())

    def _normalize(self, ranges: list[TimeRange]) -> list[TimeRange]:
        sorted_ranges = sorted(ranges, key=lambda r: r.start)
        result = []
        for r in sorted_ranges:
            if not result:
                result.append(r)
            else:
                last = result[-1]
                if last.end >= r.start:
                    merged = TimeRange(last.start, max(last.end, r.end))
                    result[-1] = merged
                else:
                    result.append(r)
        return result

    def __iter__(self):
        return iter(self.ranges)

    def __repr__(self):
        return f"TimeRangeSet({self.ranges})"


# end of file

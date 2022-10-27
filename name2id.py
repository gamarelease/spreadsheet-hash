#!/usr/bin/env python3
# vim: fileencoding=utf-8

# Project:  spreadsheet-hash
# Version:  1.0
# Created:  26-10-2022
# Authors:  Leonardo Gama
# Homepage: github.com/leogama/spreadsheet-hash

# MIT License
#
# Copyright  2022  Leonardo dos Reis Gama
#
# Permission is hereby granted, free of charge, to any person obtaining a
# copy of this software and associated documentation files (the "Software"),
# to deal in the Software without restriction, including without limitation
# the rights to use, copy, modify, merge, publish, distribute, sublicense,
# and/or sell copies of the Software, and to permit persons to whom the
# Software is furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
# THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
# FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
# DEALINGS IN THE SOFTWARE.

import re
import fnvhash

def fnv1a_32(text: str, length: int = 8, salt: str = None) -> int:
    if length not in range(1, 9):
        raise ValueError("'length' must be an integer between 1 and 8 (inclusive)")
    if salt is not None:
        text = salt + text
    hash = fnvhash.fnv1a_32(text.encode('utf8'))
    if length < 8:
        # XOR-folding
        bit_size = length*4
        size_mask = (1 << bit_size) - 1
        hash = ((hash >> bit_size) ^ hash) & size_mask
    return "{0:0{1}X}".format(hash, length)

def name2id(name: str, length: int = 8, salt: str = None, *, cast_ascii=False) -> int:
    name = name.upper().strip()
    name = name.replace('.', ' ')
    name = re.sub(r'\s+', ' ', name)
    if cast_ascii:
        import unidecode
        name = unidecode.unidecode(name)
    return fnv1a_32(name, length, salt)

if __name__ == '__main__':
    hash1 = fnv1a_32("hello world")
    hash2 = fnv1a_32("world", salt="hello ")
    print("{} == {}".format(hash1, hash2))
    hash1 = name2id("Maria  D.Assunção ", 4, cast_ascii=True)
    hash2 = name2id("MARIA D ASSUNCAO", 4)
    print("    {} == {}".format(hash1, hash2))

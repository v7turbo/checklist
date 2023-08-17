#!/usr/bin/env python

import os
from tkinter import *
from tkinter.filedialog import * 

target_list = []
try:
    with open('vhod.txt', 'r') as source_file:
        for line in source_file.readlines():
            line = line.rstrip()
            print(len(line))
            if len(line) == 12:
                line = line[:3] + '-' + line[3:6] + '-' + line[6:] + '\n'
                target_list.append(line)
            elif len(line) < 12:
                line = '0' + line[:2] + '-' + line[2:5] + '-' + line[5:] + '\n'
                target_list.append(line)
            else:
                print('Incorrect format of the entries!')
except IOError:
    print("Can't open the source. No such file!")

with open('vihod.txt', 'w') as target_file:
    for line in target_list:
        target_file.write(line)

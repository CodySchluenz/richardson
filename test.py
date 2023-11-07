#!/bin/python3

import math
import os
import random
import re
import sys



#
# Complete the 'fizzBuzz' function below.
#
# The function accepts INTEGER n as parameter.
#

def fizzBuzz(n):
    # Write your code here
    for i in range(1, n):
        buzzy = ""
        if n % 3 == 0:
            buzzy = buzzy + "Fizz"
        if n % 5 == 0:
            buzzy = buzzy + "Buzz"
        if n % 3 != 0 and n % 5 != 0:
            buzzy = n
        print(buzzy)

if __name__ == '__main__':
    
    n = int(input().strip())

    fizzBuzz(n)

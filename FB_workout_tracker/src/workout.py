import openpyxl
import math

class workout:

    def __init__(self, workbook, weight, reps):
        self.workbook = workbook
        self.weight = weight
        self.reps = reps


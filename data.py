import numpy as np
import pandas as pd

class ExcelData:
    def __init__(self, name, ratio):
        self.name = name
        self.ratio = ratio

class ExcelDatalist:
    datalist =[]

    def push(self,data):
        self.datalist.append(data)
    def printlist(self):
        for d in self.datalist:
            print(d.name,d.ratio)
    def sort_data_up(self):
        sortdatas = sorted(self.datalist, key= lambda x : x.ratio)
        self.datalist = sortdatas

    def sort_data_down(self):
        sortdatas = sorted(self.datalist, key=lambda x: x.ratio,reverse=True)
        self.datalist = sortdatas



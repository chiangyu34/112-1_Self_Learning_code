import time
import random
import os
os.chdir("D:/03Joey/高中入學管道與學習歷程檔案/小論文資料")  # Colab 換路徑使用

import openpyxl
wb = openpyxl.load_workbook('okokok.xlsx')
s1 = wb.active
# s1['B1'].value = "氣泡排序法"
# s1['c1'].value = "選擇排序法"
# s1['d1'].value = "插入排序法"
# s1['e1'].value = "快速排序法"
# s1['a2'].value = "第1次測試"
# s1['a3'].value = "第2次測試"
# s1['a4'].value = "第3次測試"
# s1['a5'].value = "第4次測試"
# s1['a6'].value = "第5次測試"
# s1['a7'].value = "平均"

for n in range(5, 6):
    # s2 = wb.copy_worksheet(s1)
    s2 = wb.active
    original_data = [random.randint(0, 10**n) for i in range(10**n)]

    def bubble(array):
        for i in range(len(array)-1, 0, -1):
            for j in range(i):
                if array[j] > array[j+1]:
                    array[j], array[j+1] = array[j+1], array[j]
                    
    def selection(array):
        for i in range(len(array)):
            min_index = i
            for j in range(i+1, len(array)):
                if array[j] < array[min_index]:
                    min_index = j
            array[i], array[min_index] = array[min_index], array[i]
                    
    def insertion(array):
        for i in range(1, len(array)):
            temp = array[i]
            j = i - 1
            while j >= 0 and temp < array[j]:
                array[j+1] = array[j]
                j -= 1
            array[j+1] = temp

    ##def shell(array):
    ##    n = len(array)
    ##    gap = n // 2 
    ##    while gap > 0: 
    ##        for i in range(gap,n): 
    ##            temp = array[i] 
    ##            j = i 
    ##            while j >= gap and array[j-gap] > temp: 
    ##                array[j] = array[j-gap] 
    ##                j -= gap 
    ##            array[j] = temp 
    ##        gap = gap // 2

    def quick(array, left_index=0, right_index=(10**n)-1):
        if left_index < right_index:
            i = left_index
            j = right_index
            pivot = array[left_index]
            while i != j:
                while array[j] > pivot and i < j:
                    j -= 1
                while array[i] <= pivot and i < j:
                    i += 1
                if i < j:
                    array[i], array[j] = array[j], array[i]
            array[left_index], array[i] = array[i], array[left_index]
            quick(array, left_index, i-1)
            quick(array, i+1, right_index)

    sorting_algorithms = [bubble, selection, insertion, quick]


    for i in range(2, 3):
        for algorithms, col in zip(sorting_algorithms, 'bcde'):         
            data = original_data.copy()
            
            t1 = time.time()
            algorithms(data)
            t2 = time.time()
            result = t2 - t1
            s2[f'{col}{i}'] = result
            wb.save("okokok.xlsx")
    else:
        for col in 'bcde':
            s2[f'{col}7'] = f"=average({col}2:{col}6)"


wb.save("okokok.xlsx")
print('finish.')
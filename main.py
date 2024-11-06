import random
import csv
import pandas as pd
import datetime
from datetime import timedelta
import calendar
from openpyxl import load_workbook
import copy




class Tframe:    # both functions make fill list of targeted time frame with the dates and the day name of the dates
    def __init__(self, datee):
        self.date = datee

    def next_30day(self):# fills the dates of the next 30 days
        for i in range(31):
            new_d = self.date + timedelta(days=i)
            main_keys_days.append(new_d.strftime("%A"))
            new_d = new_d.strftime("%d/%m/%Y")
            new_d = f"{new_d}"
            main_dic.update({new_d: []})
            main_keys.append(new_d)

    def next_month(self):# fills the varibles with the dates from the next month
        days_of_the_month = calendar.monthrange(self.date.year, self.date.month)[1]
        days_till_end_month = days_of_the_month-self.date.day
        self.date = self.date + timedelta(days=days_till_end_month+1)
        days_of_the_month = calendar.monthrange(self.date.year, self.date.month)[1]
        month = self.date.strftime('%B')
        days_name = []

        for i in range(days_of_the_month):
            new_d = self.date + timedelta(days=i)
            main_keys_days.append(new_d.strftime("%A"))
            days_name.append(new_d.strftime("%A"))
            new_d = str(int(new_d.strftime("%d")))
            new_d = f"{new_d}"
            main_dic.update({new_d: []})
            main_keys.append(new_d)



class Eframe:
    def __init__(self,main_di, main_di_v, names_list):
        self.dic_values = main_di_v.copy()
        self.WK_dic = copy.deepcopy(main_di)
        self.AM_dic = copy.deepcopy(main_di)
        self.PM_dic = copy.deepcopy(main_di)
        self.names = names_list.copy()


    def PM(self):
        global name, PM, index, selection_counts
        dicval = self.dic_values.copy()
        names_list = self.names.copy()

        index = 0  # Start from the first name
        random.shuffle(names_list)
        random.shuffle(dicval)
        random.shuffle(names_list)
        random.shuffle(names_list)

        selection_counts = {name: 0 for name in names_list}  # Initialize counts to zero

        for i in range(len(main_keys)):
            k = 0  # Counter to ensure four unique names are added per key
            while k < 4:
                # Sort names by the number of times they've been assigned, ascending
                sorted_names = sorted(selection_counts, key=selection_counts.get)

                sortlen = len(sorted_names)
                # Attempt to assign the least-used name

                for nam in range(sortlen):
                    sorted_names = sorted(selection_counts, key=selection_counts.get)
                    name = sorted_names[nam]

                    # Check for collisions across `AM_dic` and `PM_dic`
                    if name not in self.PM_dic[main_keys[i]]:
                        self.PM_dic[main_keys[i]].append(name)  # Add to AM_dic
                        selection_counts[name] += 1  # Increment the count for this name
                        k += 1  # Increment unique count for this key
                        break  # Exit inner loop to move to the next unique position


        names_list = self.names.copy()


        PM = self.PM_dic.copy()
        for e in range(len(main_keys)):
            empp = [""] * len(names_list)

            key = PM[main_keys[e]]
            for i in range(len(names_list)):
                #print(i)
                tel = key

                if names_list[i] in tel:
                     empp[i] = "PM"
            PM[main_keys[e]] = empp

    def AM(self):
        global name, AM, index
        dicval = self.dic_values.copy()
        names_list = self.names.copy()

        index = 0  # Start from the first name
        random.shuffle(names_list)
        random.shuffle(dicval)
        random.shuffle(names_list)
        random.shuffle(names_list)

        def round_robin_selection():
            global index
            if len(names_list) == 0:
                print("No names left to choose from!")
                return None

            selected_name = names_list[index]
            index = (index + 1) % len(names_list)  # Move to the next name, looping back to the start
            return selected_name


        for i in range(len(main_keys)):
            k = 0  # Counter to ensure four unique names are added per key
            while k < 4:
                # Sort names by the number of times they've been assigned, ascending
                sorted_names = sorted(selection_counts, key=selection_counts.get)

                sortlen = len(sorted_names)
                # Attempt to assign the least-used name

                for nam in range(sortlen):
                    sorted_names = sorted(selection_counts, key=selection_counts.get)
                    name = sorted_names[nam]


                    # Check for collisions across `AM_dic` and `PM_dic`
                    if name not in self.AM_dic[main_keys[i]] and name not in self.PM_dic[main_keys[i]]:
                        self.AM_dic[main_keys[i]].append(name)  # Add to AM_dic
                        selection_counts[name] += 1  # Increment the count for this name
                        k += 1  # Increment unique count for this key
                        break  # Exit inner loop to move to the next unique position
                print(name)
                print(selection_counts[name])

        names_list = self.names.copy()

        AM = self.AM_dic.copy()
        for e in range(len(main_keys)):
            empp = [""] * len(names_list)

            key = AM[main_keys[e]]
            for i in range(len(names_list)):
                # print(i)
                tel = key

                if names_list[i] in tel:
                    empp[i] = "AM"
            AM[main_keys[e]] = empp



    def print(self):
        data = {
            "names":names_list
        }

        def merge_dicts(dict1, dict2):
            merged_dict = {}

            for key in dict1.keys():
                # Get the lists from both dictionaries
                list1 = dict1[key]
                list2 = dict2[key]

                # Initialize the merged list with the same length as the original lists
                merged_list = []

                # Loop through the items in both lists
                for item1, item2 in zip(list1, list2):
                    if item1 == item2:
                        # If both items are the same, keep that item
                        merged_list.append(item1)
                    elif (item1 in ["AM", "PM"] and item2 == "") or (item2 in ["AM", "PM"] and item1 == ""):
                        # If one is AM/PM and the other is an empty string, keep the AM/PM
                        merged_list.append(item1 if item1 else item2)
                    elif (item1 == "AM" and item2 == "PM") or (item1 == "PM" and item2 == "AM"):
                        # If both lists have AM and PM, it's a collision
                        merged_list.append("collision")
                    else:
                        # If both are empty strings or other values, you can decide how to handle
                        merged_list.append("")  # Retain empty string for both empty

                merged_dict[key] = merged_list

            return merged_dict
        main_keys_days.insert(0,' ')

        #ata.update(PM)
        #data.update(AM)
        data.update(merge_dicts(PM, AM))
        df = pd.DataFrame(data)
        df.to_excel("data.xlsx", index=False, engine='openpyxl')
        wb = load_workbook("data.xlsx")  # Replace with your file name
        ws = wb.active
        ws.insert_rows(1)
        for col_num, value in enumerate(main_keys_days, start=1):  # start=1 to start from column A
            ws.cell(row=1, column=col_num, value=value[0])
        wb.save("data.xlsx")


    def count_shifts(self):
        for wq in range(len(self.names)):
            namess = self.names[wq]
            s = 0
            for d in range(len(self.dic_values)):
                rand = self.dic_values[d]
                if namess in self.PM_dic[rand]:
                    s += 1
                if namess in self.AM_dic[rand]:
                    s += 1




            if s == 8:
                print(f"{namess}={s}5888888888888885")
            else:
                print(f"{namess}={s}")

    def count_days_shifts(self):
        for ein in range(len(self.dic_values)):
            print(f"{self.dic_values[ein]}={len(self.AM_dic[self.dic_values[ein]])+len(self.PM_dic[self.dic_values[ein]])}")

    def count_emps(self):
        namelist = self.names.copy()
        print(f"we have {len(namelist)} avalible employees")






main_dic = {}  # each key is a date and contains the names of people who are working on these days
main_keys = []  # a list of each key easier to handle
main_keys_days = [] # names of the dates like m for monday
names_list = [] # to be filled with names of the csv file to handle it easier


with open('names.csv', 'r') as na:# fills names in a list
    names = csv.reader(na)
    next(names)# skips the first line in the file
    for lk in names:
        names_list.append(lk[0])

date = datetime.datetime.now()

T = Tframe(date)  # T will fill the necessary list to be able to distribute emps shifts
T.next_month()
e = Eframe(main_dic, main_keys, names_list)
e.PM()
e.AM()
e.print()

e.count_shifts()
e.count_days_shifts()

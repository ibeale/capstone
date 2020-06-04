import random
from openpyxl import Workbook
import json

# https://www.itl.nist.gov/div898/handbook/pri/section3/pri3347.htm
def main():
    input("Warning: This will delete the previous TestingOrder.xlsx file. Stop this program now (ctrl+c) and rename the old TestingOrder.xlsx file if you would like to save its contents. Press ENTER to continue.")
    wb = Workbook()
    ws = wb.active
    ws.title = "Testing Order"
    finalList = []

    with open("key.csv", "r") as f:
        keylines = f.read().splitlines()
        keyheaders = keylines[0].split(',')[1:]
        ws.append(keyheaders)
        pos = keylines[1].split(',')
        neg = keylines[2].split(',')
        keyDict = {}
        for i in range(1,len(keyheaders)+1):
            keyDict[keyheaders[i-1]] = {}
            keyDict[keyheaders[i-1]][pos[i]] = pos[0]
            keyDict[keyheaders[i-1]][neg[i]] = neg[0]
        with open("key.json", "w") as f:
            json.dump(keyDict, f)
        pos = pos[1:]
        neg = neg[1:]

    with open("states.txt", "r") as f:
        stateslines = f.read().splitlines()

    for line in stateslines:
        stateList = line.split()
        if len(stateList) != len(keyheaders):
            print(f"The number factors in the key.csv file ({len(keyheaders)}) does not match the number of factors in the states.txt file ({len(stateList)}).")
            return
        for i in range(len(stateList)):
            if stateList[i] == '+1':
                stateList[i] = pos[i]
            else:
                stateList[i] = neg[i]
        finalList.append(stateList)
    random.shuffle(finalList)
    for item in finalList:
        ws.append(item)
            
    wb.save(filename="TestingOrder.xlsx")

if __name__ == "__main__":
    main()
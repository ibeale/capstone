import random


def main():
    finalList = []
    with open("states.txt", "r") as f:
        list_of_strings = f.read().splitlines()
    with open("out.csv", "w") as o:
        for item in list_of_strings:
            finalList.append(','.join(item.split()) + '\n')
        random.shuffle(finalList)
        for item in finalList:
            o.write(item)


if __name__ == "__main__":
    main()
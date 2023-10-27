import os
import csv
file = open('Cartons.csv')
reader = csv.reader(file)
y = 1
for x in reader:
    if y == 1:
        y = 2
        continue
    parent_dir = 'Carton-Specifications'
    directory = str(x[0])
    path = os.path.join(parent_dir,directory)
    os.mkdir(path)
    print("Directory '% s' created" % directory)

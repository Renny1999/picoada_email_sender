import csv

data = []
with open('emails.csv','r') as csvfile:
  reader = csv.reader(csvfile, delimiter=',', quotechar='\"')
  i = 0
  for row in reader:
    data.append(row)


print(data[37][25])
